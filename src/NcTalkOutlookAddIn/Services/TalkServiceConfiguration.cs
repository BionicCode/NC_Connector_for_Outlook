// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.Services
{
        // Encapsulates connection details required for Nextcloud Talk REST calls.
    internal sealed class TalkServiceConfiguration
    {
        internal string BaseUrl { get; private set; }
        internal string Username { get; private set; }
        internal string AppPassword { get; private set; }

        public TalkServiceConfiguration(string baseUrl, string username, string appPassword)
        {
            BaseUrl = baseUrl ?? string.Empty;
            Username = username ?? string.Empty;
            AppPassword = appPassword ?? string.Empty;
        }

        // Returns a canonical HTTPS base URL or an empty value for an untrusted URL.
        public string GetNormalizedBaseUrl()
        {
            string normalized;
            return NextcloudUriValidator.TryNormalizeBaseUrl(BaseUrl, out normalized)
                ? normalized
                : string.Empty;
        }

        public bool IsComplete()
        {
            return !string.IsNullOrWhiteSpace(GetNormalizedBaseUrl())
                   && !string.IsNullOrWhiteSpace(Username)
                   && !string.IsNullOrWhiteSpace(AppPassword);
        }
    }
}
