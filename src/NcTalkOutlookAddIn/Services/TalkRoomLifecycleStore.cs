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
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.Services
{
    internal sealed class TalkRoomLifecycleStore
    {
        private static readonly byte[] ProtectionEntropy =
            Encoding.UTF8.GetBytes("NC4OL::TalkRoomLifecycle::v1");

        private readonly object _syncRoot = new object();
        private readonly string _filePath;
        private readonly string _backupFilePath;
        private readonly JavaScriptSerializer _serializer =
            new JavaScriptSerializer();
        private bool _writeAllowed = true;

        internal TalkRoomLifecycleStore(
            string dataDirectory,
            string profileScope)
        {
            string directory = string.IsNullOrWhiteSpace(dataDirectory)
                ? AppDataPaths.EnsureLocalRootDirectory()
                : dataDirectory;
            Directory.CreateDirectory(directory);
            _filePath = Path.Combine(
                directory,
                "talk-room-lifecycle-"
                + BuildScopeHash(profileScope)
                + ".dat");
            _backupFilePath = _filePath + ".bak";
        }

        internal TalkRoomLifecycleState Load()
        {
            lock (_syncRoot)
            {
                TalkRoomLifecycleState state;
                if (TryLoadFile(_filePath, out state))
                {
                    return state;
                }
                if (TryLoadFile(_backupFilePath, out state))
                {
                    DiagnosticsLogger.Log(
                        LogCategories.Talk,
                        "Recovered the Talk room deletion queue from its backup.");
                    try
                    {
                        File.Copy(
                            _backupFilePath,
                            _filePath,
                            true);
                    }
                    catch (Exception ex)
                    {
                        DiagnosticsLogger.LogException(
                            LogCategories.Talk,
                            "Failed to restore the Talk room deletion queue from its backup.",
                            ex);
                    }
                    return state;
                }

                if (File.Exists(_filePath)
                    || File.Exists(_backupFilePath))
                {
                    _writeAllowed = false;
                    DiagnosticsLogger.Log(
                        LogCategories.Talk,
                        "Talk room deletion queue recovery failed; writes are disabled to preserve the existing files.");
                }
                return new TalkRoomLifecycleState();
            }
        }

        internal void Save(TalkRoomLifecycleState state)
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
                        "Talk room deletion queue is unreadable; existing data was preserved.");
                }
                string json = _serializer.Serialize(state);
                byte[] protectedBytes = ProtectedData.Protect(
                    Encoding.UTF8.GetBytes(json),
                    ProtectionEntropy,
                    DataProtectionScope.CurrentUser);
                string encoded = Convert.ToBase64String(protectedBytes);
                string temporaryPath =
                    _filePath + "." + Guid.NewGuid().ToString("N") + ".tmp";
                try
                {
                    File.WriteAllText(
                        temporaryPath,
                        encoded,
                        new UTF8Encoding(false));
                    if (DurableFileReplace.CommitPreparedFile(
                            temporaryPath,
                            _filePath,
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
            out TalkRoomLifecycleState state)
        {
            state = null;
            if (!File.Exists(path))
            {
                return false;
            }

            try
            {
                string encoded = File.ReadAllText(path, Encoding.UTF8);
                byte[] protectedBytes = Convert.FromBase64String(encoded);
                byte[] clearBytes = ProtectedData.Unprotect(
                    protectedBytes,
                    ProtectionEntropy,
                    DataProtectionScope.CurrentUser);
                TalkRoomLifecycleState candidate =
                    _serializer.Deserialize<TalkRoomLifecycleState>(
                        Encoding.UTF8.GetString(clearBytes));
                if (!IsValid(candidate))
                {
                    throw new InvalidDataException(
                        "Talk room deletion queue structure is invalid.");
                }
                state = candidate;
                return true;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(
                    LogCategories.Talk,
                    "Failed to load Talk room deletion queue file '"
                    + Path.GetFileName(path)
                    + "'.",
                    ex);
                return false;
            }
        }

        private static bool IsValid(TalkRoomLifecycleState state)
        {
            if (state == null || state.Records == null)
            {
                return false;
            }
            var ids = new HashSet<string>(StringComparer.Ordinal);
            for (int i = 0; i < state.Records.Count; i++)
            {
                TalkRoomLifecycleRecord record = state.Records[i];
                if (record == null
                    || string.IsNullOrWhiteSpace(record.Id)
                    || string.IsNullOrWhiteSpace(record.RoomToken)
                    || string.IsNullOrWhiteSpace(record.ServerBaseUrl)
                    || !ids.Add(record.Id))
                {
                    return false;
                }
            }
            return true;
        }

        private static string BuildScopeHash(string profileScope)
        {
            string normalized = string.IsNullOrWhiteSpace(profileScope)
                ? "default"
                : profileScope.Trim().ToUpperInvariant();
            byte[] hash;
            using (SHA256 sha256 = SHA256.Create())
            {
                hash = sha256.ComputeHash(
                    Encoding.UTF8.GetBytes(normalized));
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

    internal sealed class TalkRoomLifecycleState
    {
        public TalkRoomLifecycleState()
        {
            Records = new List<TalkRoomLifecycleRecord>();
        }

        public List<TalkRoomLifecycleRecord> Records { get; set; }
    }
}
