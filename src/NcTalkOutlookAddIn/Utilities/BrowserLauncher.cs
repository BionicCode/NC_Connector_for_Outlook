// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Diagnostics;

namespace NcTalkOutlookAddIn.Utilities
{
        // Central utility for opening shell targets (URLs, files, folders).
    internal static class BrowserLauncher
    {
        internal static bool OpenTarget(string target, string logCategory, string failureContext)
        {
            string unused;
            return OpenTarget(target, logCategory, failureContext, out unused);
        }

        internal static bool OpenTarget(string target, string logCategory, string failureContext, out string errorMessage)
        {
            errorMessage = string.Empty;
            if (string.IsNullOrWhiteSpace(target))
            {
                errorMessage = "Target is empty.";
                return false;
            }
            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = target,
                    UseShellExecute = true
                });
                return true;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(logCategory, failureContext ?? "Failed to open shell target.", ex);
                errorMessage = ex.Message ?? string.Empty;
                return false;
            }
        }

        internal static bool OpenUrl(string url, string logCategory, string failureContext)
        {
            string normalizedUrl;
            if (!NextcloudUriValidator.TryNormalizeHttpsUrl(url, out normalizedUrl))
            {
                DiagnosticsLogger.Log(
                    logCategory,
                    (failureContext ?? "Failed to open URL.") + " URL rejected because it is not a valid HTTPS URL.");
                return false;
            }

            return OpenTarget(normalizedUrl, logCategory, failureContext ?? "Failed to open URL.");
        }
    }
}
