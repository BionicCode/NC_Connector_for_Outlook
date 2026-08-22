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
    // Caches system-address-book email and UID mappings for one Outlook profile.
    internal sealed class IfbAddressBookCache
    {
        private readonly object _syncRoot = new object();
        private readonly string _dataDirectory;
        private readonly string _profileScope;
        private readonly JavaScriptSerializer _serializer =
            new JavaScriptSerializer();

        private Dictionary<string, string> _emailToUid =
            new Dictionary<string, string>(
                StringComparer.OrdinalIgnoreCase);
        private Dictionary<string, string> _uidToEmail =
            new Dictionary<string, string>(
                StringComparer.OrdinalIgnoreCase);
        private DateTime _generatedUtc = DateTime.MinValue;
        private string _activeScopeFingerprint = string.Empty;

        internal IfbAddressBookCache(string dataDirectory)
            : this(dataDirectory, "default")
        {
        }

        internal IfbAddressBookCache(
            string dataDirectory,
            string profileScope)
        {
            _dataDirectory = string.IsNullOrWhiteSpace(dataDirectory)
                ? AppDataPaths.EnsureLocalRootDirectory()
                : dataDirectory;
            _profileScope = string.IsNullOrWhiteSpace(profileScope)
                ? "default"
                : profileScope.Trim();
            Directory.CreateDirectory(_dataDirectory);
        }

        internal sealed class SystemAddressbookStatus
        {
            internal SystemAddressbookStatus(
                bool available,
                int count,
                string error)
            {
                Available = available;
                Count = count;
                Error = error ?? string.Empty;
            }

            internal bool Available { get; private set; }

            internal int Count { get; private set; }

            internal string Error { get; private set; }
        }

        internal SystemAddressbookStatus GetSystemAddressbookStatus(
            TalkServiceConfiguration configuration,
            int cacheHours,
            bool forceRefresh)
        {
            if (configuration == null
                || !configuration.IsComplete())
            {
                const string detail =
                    "Talk credentials are incomplete.";
                DiagnosticsLogger.Log(
                    LogCategories.Ifb,
                    "System address book status check failed: "
                    + detail);
                return new SystemAddressbookStatus(
                    false,
                    0,
                    detail);
            }

            lock (_syncRoot)
            {
                try
                {
                    CacheScope scope = CreateScope(configuration);
                    if (forceRefresh)
                    {
                        RefreshFromServer(
                            configuration,
                            scope);
                    }
                    else
                    {
                        EnsureCache(
                            configuration,
                            cacheHours,
                            scope);
                    }
                    return new SystemAddressbookStatus(
                        true,
                        _uidToEmail.Count,
                        string.Empty);
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(
                        LogCategories.Ifb,
                        "System address book status check failed.",
                        ex);
                    return new SystemAddressbookStatus(
                        false,
                        0,
                        ex.Message
                        ?? "System address book status check failed.");
                }
            }
        }

        internal bool TryGetUid(
            TalkServiceConfiguration configuration,
            int cacheHours,
            string email,
            out string uid)
        {
            uid = null;
            if (configuration == null
                || !configuration.IsComplete()
                || string.IsNullOrWhiteSpace(email))
            {
                return false;
            }

            lock (_syncRoot)
            {
                EnsureCache(
                    configuration,
                    cacheHours,
                    CreateScope(configuration));
                return _emailToUid.TryGetValue(
                    email.Trim().ToLowerInvariant(),
                    out uid);
            }
        }

        internal bool TryGetPrimaryEmailForUid(
            TalkServiceConfiguration configuration,
            int cacheHours,
            string uid,
            out string email)
        {
            email = null;
            if (configuration == null
                || !configuration.IsComplete()
                || string.IsNullOrWhiteSpace(uid))
            {
                return false;
            }

            lock (_syncRoot)
            {
                EnsureCache(
                    configuration,
                    cacheHours,
                    CreateScope(configuration));
                return _uidToEmail.TryGetValue(
                    uid.Trim(),
                    out email);
            }
        }

        internal List<NextcloudUser> GetUsers(
            TalkServiceConfiguration configuration,
            int cacheHours,
            bool forceRefresh)
        {
            var users = new List<NextcloudUser>();
            if (configuration == null
                || !configuration.IsComplete())
            {
                return users;
            }

            lock (_syncRoot)
            {
                CacheScope scope = CreateScope(configuration);
                if (forceRefresh)
                {
                    RefreshFromServer(
                        configuration,
                        scope);
                }
                else
                {
                    EnsureCache(
                        configuration,
                        cacheHours,
                        scope);
                }
                foreach (KeyValuePair<string, string> pair
                    in _uidToEmail)
                {
                    if (!string.IsNullOrWhiteSpace(pair.Key))
                    {
                        users.Add(
                            new NextcloudUser(
                                pair.Key.Trim(),
                                pair.Value
                                ?? string.Empty));
                    }
                }
            }

            users.Sort(
                (left, right) => string.Compare(
                    left.UserId,
                    right.UserId,
                    StringComparison.OrdinalIgnoreCase));
            return users;
        }

        private void EnsureCache(
            TalkServiceConfiguration configuration,
            int cacheHours,
            CacheScope scope)
        {
            int validHours = Math.Max(1, cacheHours);
            if (string.Equals(
                    _activeScopeFingerprint,
                    scope.Fingerprint,
                    StringComparison.Ordinal)
                && _generatedUtc > DateTime.MinValue
                && _generatedUtc.AddHours(validHours)
                   > DateTime.UtcNow)
            {
                return;
            }

            // Scope construction is local, so disk cache is tried before UID resolution.
            if (!LoadFromDisk(scope, validHours))
            {
                RefreshFromServer(configuration, scope);
            }
        }

        private bool LoadFromDisk(
            CacheScope scope,
            int cacheHours)
        {
            string path = BuildCacheFilePath(scope);
            if (!File.Exists(path))
            {
                return false;
            }

            try
            {
                CacheContainer data =
                    _serializer.Deserialize<CacheContainer>(
                        File.ReadAllText(path, Encoding.UTF8));
                if (data == null
                    || data.GeneratedUtc <= DateTime.MinValue
                    || data.Entries == null
                    || data.GeneratedUtc.AddHours(cacheHours)
                       <= DateTime.UtcNow
                    || !string.Equals(
                        data.ScopeFingerprint,
                        scope.Fingerprint,
                        StringComparison.Ordinal))
                {
                    return false;
                }

                ApplyEntries(
                    data.Entries,
                    data.GeneratedUtc,
                    scope.Fingerprint);
                return true;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(
                    LogCategories.Ifb,
                    "Failed to load IFB address book cache from disk.",
                    ex);
                return false;
            }
        }

        private void RefreshFromServer(
            TalkServiceConfiguration configuration,
            CacheScope scope)
        {
            string currentUserId =
                NextcloudUserIdentityService.ResolveCurrentUserId(
                    configuration);
            string addressBookUrl = string.Format(
                CultureInfo.InvariantCulture,
                "{0}/remote.php/dav/addressbooks/users/{1}/z-server-generated--system?export",
                scope.ServerBaseUrl,
                Uri.EscapeDataString(currentUserId));

            var httpClient = new NcHttpClient(configuration);
            NcHttpResponse response = httpClient.Send(
                new NcHttpRequestOptions
                {
                    Method = "GET",
                    Url = addressBookUrl,
                    Accept =
                        "text/vcard,text/x-vcard,text/plain,*/*",
                    TimeoutMs = 60000,
                    IncludeAuthHeader = true,
                    IncludeOcsApiHeader = false,
                    ParseJson = false
                });
            if (!response.HasHttpResponse)
            {
                if (response.TransportException != null)
                {
                    throw new InvalidOperationException(
                        "Address book could not be loaded: "
                        + response.TransportException.Message,
                        response.TransportException);
                }
                throw new InvalidOperationException(
                    "Address book could not be loaded: no HTTP response.");
            }
            if ((int)response.StatusCode < 200
                || (int)response.StatusCode >= 300)
            {
                throw new InvalidOperationException(
                    "Address book could not be loaded: HTTP "
                    + ((int)response.StatusCode).ToString(
                        CultureInfo.InvariantCulture)
                    + ".");
            }
            string responseText =
                response.ResponseText ?? string.Empty;
            if (responseText.Length == 0)
            {
                throw new InvalidOperationException(
                    "Address book response was empty.");
            }

            List<CacheEntry> entries =
                ParseAddressBook(responseText);
            DateTime generatedUtc = DateTime.UtcNow;
            ApplyEntries(
                entries,
                generatedUtc,
                scope.Fingerprint);
            SaveToDisk(scope, generatedUtc);
        }

        private void ApplyEntries(
            IEnumerable<CacheEntry> entries,
            DateTime generatedUtc,
            string scopeFingerprint)
        {
            var emailMap = new Dictionary<string, string>(
                StringComparer.OrdinalIgnoreCase);
            var uidMap = new Dictionary<string, string>(
                StringComparer.OrdinalIgnoreCase);
            foreach (CacheEntry entry in entries)
            {
                if (entry == null
                    || string.IsNullOrWhiteSpace(entry.Email)
                    || string.IsNullOrWhiteSpace(entry.Uid))
                {
                    continue;
                }
                string email =
                    entry.Email.Trim().ToLowerInvariant();
                string uid = entry.Uid.Trim();
                if (!emailMap.ContainsKey(email))
                {
                    emailMap[email] = uid;
                }
                if (!uidMap.ContainsKey(uid))
                {
                    uidMap[uid] = email;
                }
            }

            _emailToUid = emailMap;
            _uidToEmail = uidMap;
            _generatedUtc = generatedUtc;
            _activeScopeFingerprint = scopeFingerprint;
        }

        private void SaveToDisk(
            CacheScope scope,
            DateTime generatedUtc)
        {
            var data = new CacheContainer
            {
                GeneratedUtc = generatedUtc,
                ScopeFingerprint = scope.Fingerprint,
                Entries = new List<CacheEntry>()
            };
            foreach (KeyValuePair<string, string> pair
                in _emailToUid)
            {
                data.Entries.Add(
                    new CacheEntry
                    {
                        Email = pair.Key,
                        Uid = pair.Value
                    });
            }

            try
            {
                File.WriteAllText(
                    BuildCacheFilePath(scope),
                    _serializer.Serialize(data),
                    new UTF8Encoding(false));
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(
                    LogCategories.Ifb,
                    "Failed to write IFB address book cache to disk.",
                    ex);
            }
        }

        private CacheScope CreateScope(
            TalkServiceConfiguration configuration)
        {
            string serverBaseUrl =
                configuration.GetNormalizedBaseUrl();
            if (string.IsNullOrWhiteSpace(serverBaseUrl))
            {
                throw new InvalidOperationException(
                    "Server URL is invalid.");
            }
            string username =
                (configuration.Username
                 ?? string.Empty).Trim();
            return new CacheScope(
                serverBaseUrl,
                BuildScopeFingerprint(
                    _profileScope,
                    serverBaseUrl,
                    username));
        }

        private string BuildCacheFilePath(CacheScope scope)
        {
            return Path.Combine(
                _dataDirectory,
                "ifb-addressbook-cache-"
                + scope.Fingerprint
                + ".json");
        }

        internal static string BuildScopeFingerprint(
            string profileScope,
            string serverBaseUrl,
            string username)
        {
            string input =
                (profileScope ?? string.Empty).Trim()
                + "\n"
                + (serverBaseUrl ?? string.Empty).TrimEnd('/')
                + "\n"
                + (username ?? string.Empty).Trim();
            byte[] hash;
            using (SHA256 sha256 = SHA256.Create())
            {
                hash = sha256.ComputeHash(
                    Encoding.UTF8.GetBytes(input));
            }

            var builder = new StringBuilder(32);
            for (int i = 0; i < 16; i++)
            {
                builder.Append(
                    hash[i].ToString(
                        "x2",
                        CultureInfo.InvariantCulture));
            }
            return builder.ToString();
        }

        private static List<CacheEntry> ParseAddressBook(
            string data)
        {
            var result = new List<CacheEntry>();
            string normalized = (data ?? string.Empty)
                .Replace("\r\n ", string.Empty)
                .Replace("\n ", string.Empty)
                .Replace("\r\n\t", string.Empty)
                .Replace("\n\t", string.Empty);
            using (var reader = new StringReader(normalized))
            {
                string line;
                string uid = null;
                var emails = new List<string>();
                bool inside = false;
                while ((line = reader.ReadLine()) != null)
                {
                    if (line.StartsWith(
                            "BEGIN:VCARD",
                            StringComparison.OrdinalIgnoreCase))
                    {
                        inside = true;
                        uid = null;
                        emails.Clear();
                    }
                    else if (line.StartsWith(
                                 "END:VCARD",
                                 StringComparison.OrdinalIgnoreCase))
                    {
                        if (inside
                            && !string.IsNullOrWhiteSpace(uid))
                        {
                            foreach (string email in emails)
                            {
                                result.Add(
                                    new CacheEntry
                                    {
                                        Email = email,
                                        Uid = uid.Trim()
                                    });
                            }
                        }
                        inside = false;
                    }
                    else if (inside
                             && line.StartsWith(
                                 "UID",
                                 StringComparison.OrdinalIgnoreCase))
                    {
                        uid = ReadVCardValue(line, 3);
                    }
                    else if (inside
                             && line.StartsWith(
                                 "EMAIL",
                                 StringComparison.OrdinalIgnoreCase))
                    {
                        string email =
                            ReadVCardValue(line, 5)
                                .ToLowerInvariant();
                        if (email.Length > 0)
                        {
                            emails.Add(email);
                        }
                    }
                }
            }
            return result;
        }

        private static string ReadVCardValue(
            string line,
            int searchStart)
        {
            int colon = line.IndexOf(':', searchStart);
            return colon >= 0 && colon + 1 < line.Length
                ? line.Substring(colon + 1).Trim()
                : string.Empty;
        }

        private sealed class CacheScope
        {
            internal CacheScope(
                string serverBaseUrl,
                string fingerprint)
            {
                ServerBaseUrl = serverBaseUrl;
                Fingerprint = fingerprint;
            }

            internal string ServerBaseUrl { get; private set; }

            internal string Fingerprint { get; private set; }
        }

        private sealed class CacheContainer
        {
            public DateTime GeneratedUtc { get; set; }

            public string ScopeFingerprint { get; set; }

            public List<CacheEntry> Entries { get; set; }
        }

        private sealed class CacheEntry
        {
            public string Email { get; set; }

            public string Uid { get; set; }
        }
    }
}
