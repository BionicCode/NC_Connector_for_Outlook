// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;

namespace NcTalkOutlookAddIn.Utilities
{
    /// <summary>
    /// Central trust-boundary validation for Nextcloud URLs.
    /// Authentication data may only be sent to HTTPS endpoints on the configured origin.
    /// </summary>
    internal static class NextcloudUriValidator
    {
        internal static bool TryNormalizeBaseUrl(string rawUrl, out string normalizedUrl)
        {
            normalizedUrl = string.Empty;
            string candidate = rawUrl == null ? string.Empty : rawUrl.Trim();
            if (candidate.Length == 0 || ContainsControlCharacters(candidate))
            {
                return false;
            }

            if (candidate.IndexOf("://", StringComparison.Ordinal) < 0)
            {
                candidate = "https://" + candidate;
            }

            Uri uri;
            if (!TryCreateTrustedHttpsUri(candidate, out uri)
                || !string.IsNullOrEmpty(uri.UserInfo)
                || !string.IsNullOrEmpty(uri.Query)
                || !string.IsNullOrEmpty(uri.Fragment))
            {
                return false;
            }

            normalizedUrl = uri.GetLeftPart(UriPartial.Path).TrimEnd('/');
            return normalizedUrl.Length > 0;
        }

        internal static bool TryResolveSameOriginHttpsUrl(
            string rawUrl,
            string baseUrl,
            out string resolvedUrl)
        {
            resolvedUrl = string.Empty;

            string normalizedBase;
            if (!TryNormalizeBaseUrl(baseUrl, out normalizedBase))
            {
                return false;
            }

            string candidate = rawUrl == null ? string.Empty : rawUrl.Trim();
            if (candidate.Length == 0 || ContainsControlCharacters(candidate))
            {
                return false;
            }

            Uri baseUri;
            Uri resolvedUri;
            if (!Uri.TryCreate(normalizedBase + "/", UriKind.Absolute, out baseUri)
                || !Uri.TryCreate(baseUri, candidate, out resolvedUri)
                || !IsTrustedHttpsUri(resolvedUri)
                || !HasSameOrigin(baseUri, resolvedUri))
            {
                return false;
            }

            resolvedUrl = resolvedUri.AbsoluteUri;
            return true;
        }

        internal static bool TryNormalizeHttpsUrl(string rawUrl, out string normalizedUrl)
        {
            normalizedUrl = string.Empty;
            string candidate = rawUrl == null ? string.Empty : rawUrl.Trim();
            if (candidate.Length == 0 || ContainsControlCharacters(candidate))
            {
                return false;
            }

            Uri uri;
            if (!TryCreateTrustedHttpsUri(candidate, out uri))
            {
                return false;
            }

            normalizedUrl = uri.AbsoluteUri;
            return true;
        }

        internal static bool HasSameOrigin(Uri expectedOrigin, Uri candidate)
        {
            return expectedOrigin != null
                && candidate != null
                && string.Equals(expectedOrigin.Scheme, candidate.Scheme, StringComparison.OrdinalIgnoreCase)
                && string.Equals(expectedOrigin.IdnHost, candidate.IdnHost, StringComparison.OrdinalIgnoreCase)
                && expectedOrigin.Port == candidate.Port;
        }

        private static bool TryCreateTrustedHttpsUri(string candidate, out Uri uri)
        {
            uri = null;
            Uri parsed;
            if (!Uri.TryCreate(candidate, UriKind.Absolute, out parsed) || !IsTrustedHttpsUri(parsed))
            {
                return false;
            }

            uri = parsed;
            return true;
        }

        private static bool IsTrustedHttpsUri(Uri uri)
        {
            return uri != null
                && uri.IsAbsoluteUri
                && string.Equals(uri.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase)
                && !string.IsNullOrWhiteSpace(uri.Host)
                && string.IsNullOrEmpty(uri.UserInfo);
        }

        private static bool ContainsControlCharacters(string value)
        {
            for (int i = 0; i < value.Length; i++)
            {
                if (char.IsControl(value[i]))
                {
                    return true;
                }
            }
            return false;
        }
    }
}
