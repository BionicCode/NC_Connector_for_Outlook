// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Web.Script.Serialization;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.Services
{
    internal sealed class IfbRegistryStateStore
    {
        private static readonly byte[] ProtectionEntropy =
            Encoding.UTF8.GetBytes("NC4OL::IFB::RegistryState::v1");

        private readonly object _syncRoot = new object();
        private readonly string _stateFilePath;
        private readonly string _backupFilePath;
        private readonly JavaScriptSerializer _serializer =
            new JavaScriptSerializer();
        private bool _writeAllowed = true;

        internal IfbRegistryStateStore(
            string dataDirectory,
            string profileScope)
        {
            string directory = string.IsNullOrWhiteSpace(dataDirectory)
                ? AppDataPaths.EnsureLocalRootDirectory()
                : dataDirectory;
            Directory.CreateDirectory(directory);
            _stateFilePath = Path.Combine(
                directory,
                "ifb-registry-state-"
                + BuildScopeHash(profileScope)
                + ".dat");
            _backupFilePath = _stateFilePath + ".bak";
        }

        internal IfbRegistryState Load()
        {
            lock (_syncRoot)
            {
                IfbRegistryState state;
                if (TryLoadFile(_stateFilePath, out state))
                {
                    return state;
                }
                if (TryLoadFile(_backupFilePath, out state))
                {
                    DiagnosticsLogger.Log(
                        LogCategories.Ifb,
                        "Recovered the IFB registry state from its backup.");
                    try
                    {
                        File.Copy(
                            _backupFilePath,
                            _stateFilePath,
                            true);
                    }
                    catch (Exception ex)
                    {
                        DiagnosticsLogger.LogException(
                            LogCategories.Ifb,
                            "Failed to restore the IFB registry state backup.",
                            ex);
                    }
                    return state;
                }
                if (File.Exists(_stateFilePath)
                    || File.Exists(_backupFilePath))
                {
                    _writeAllowed = false;
                    DiagnosticsLogger.Log(
                        LogCategories.Ifb,
                        "IFB registry state recovery failed; the existing files were preserved.");
                }
                return new IfbRegistryState();
            }
        }

        internal void Save(IfbRegistryState state)
        {
            if (state == null)
            {
                throw new ArgumentNullException("state");
            }

            lock (_syncRoot)
            {
                if (!_writeAllowed)
                {
                    throw new InvalidOperationException(
                        "IFB registry state is unreadable; existing data was preserved.");
                }

                byte[] clearBytes = Encoding.UTF8.GetBytes(
                    _serializer.Serialize(state));
                byte[] protectedBytes = ProtectedData.Protect(
                    clearBytes,
                    ProtectionEntropy,
                    DataProtectionScope.CurrentUser);
                string temporaryPath =
                    _stateFilePath
                    + "."
                    + Guid.NewGuid().ToString("N")
                    + ".tmp";
                try
                {
                    File.WriteAllText(
                        temporaryPath,
                        Convert.ToBase64String(protectedBytes),
                        new UTF8Encoding(false));
                    if (DurableFileReplace.CommitPreparedFile(
                            temporaryPath,
                            _stateFilePath,
                            _backupFilePath))
                    {
                        temporaryPath = null;
                    }
                }
                finally
                {
                    if (!string.IsNullOrEmpty(temporaryPath)
                        && File.Exists(temporaryPath))
                    {
                        File.Delete(temporaryPath);
                    }
                }
            }
        }

        private bool TryLoadFile(
            string path,
            out IfbRegistryState state)
        {
            state = null;
            if (!File.Exists(path))
            {
                return false;
            }

            try
            {
                byte[] protectedBytes = Convert.FromBase64String(
                    File.ReadAllText(path, Encoding.UTF8));
                byte[] clearBytes = ProtectedData.Unprotect(
                    protectedBytes,
                    ProtectionEntropy,
                    DataProtectionScope.CurrentUser);
                IfbRegistryState candidate =
                    _serializer.Deserialize<IfbRegistryState>(
                        Encoding.UTF8.GetString(clearBytes));
                if (!IsValid(candidate))
                {
                    throw new InvalidDataException(
                        "IFB registry state structure is invalid.");
                }
                state = candidate;
                return true;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(
                    LogCategories.Ifb,
                    "Failed to load IFB registry state file '"
                    + Path.GetFileName(path)
                    + "'.",
                    ex);
                return false;
            }
        }

        private static bool IsValid(IfbRegistryState state)
        {
            if (state == null
                || state.Ownership == null
                || state.Ownership.Count > 2)
            {
                return false;
            }

            var keys = new HashSet<string>(
                StringComparer.OrdinalIgnoreCase);
            foreach (IfbRegistryOwnership ownership in state.Ownership)
            {
                if (ownership == null
                    || string.IsNullOrWhiteSpace(ownership.RegistryPath)
                    || string.IsNullOrWhiteSpace(ownership.ValueName)
                    || string.IsNullOrWhiteSpace(ownership.WrittenValue)
                    || !keys.Add(
                        ownership.RegistryPath
                        + "\n"
                        + ownership.ValueName))
                {
                    return false;
                }
            }
            return true;
        }

        private static string BuildScopeHash(string profileScope)
        {
            byte[] input = Encoding.UTF8.GetBytes(
                string.IsNullOrWhiteSpace(profileScope)
                    ? "default"
                    : profileScope.Trim().ToUpperInvariant());
            byte[] hash;
            using (SHA256 sha256 = SHA256.Create())
            {
                hash = sha256.ComputeHash(input);
            }

            var builder = new StringBuilder(24);
            for (int i = 0; i < 12; i++)
            {
                builder.Append(
                    hash[i].ToString(
                        "x2",
                        CultureInfo.InvariantCulture));
            }
            return builder.ToString();
        }
    }

    internal sealed class IfbRegistryState
    {
        public IfbRegistryState()
        {
            Ownership = new List<IfbRegistryOwnership>();
        }

        public List<IfbRegistryOwnership> Ownership { get; set; }
    }

    internal sealed class IfbRegistryOwnership
    {
        public string RegistryPath { get; set; }

        public string ValueName { get; set; }

        public bool OriginalExists { get; set; }

        public int OriginalKind { get; set; }

        public string OriginalValue { get; set; }

        public string WrittenValue { get; set; }
    }
}
