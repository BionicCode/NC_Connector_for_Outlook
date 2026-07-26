// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using NcTalkOutlookAddIn.Models;

namespace NcTalkOutlookAddIn.Controllers
{
    // Tracks shares that Outlook has not written with the compose item yet.
    internal sealed class ComposeShareCleanupTracker
    {
        private readonly List<ComposeShareCleanupRecord> _records =
            new List<ComposeShareCleanupRecord>();

        internal int Count
        {
            get { return _records.Count; }
        }

        internal bool Arm(ComposeShareCleanupRecord record)
        {
            if (record == null
                || string.IsNullOrWhiteSpace(record.RelativeFolder))
            {
                return false;
            }

            for (int i = 0; i < _records.Count; i++)
            {
                if (IsSameShare(_records[i], record))
                {
                    return false;
                }
            }

            _records.Add(record);
            return true;
        }

        internal int ReleaseAll()
        {
            int count = _records.Count;
            _records.Clear();
            return count;
        }

        internal List<ComposeShareCleanupRecord> Drain()
        {
            var records =
                new List<ComposeShareCleanupRecord>(_records);
            _records.Clear();
            return records;
        }

        private static bool IsSameShare(
            ComposeShareCleanupRecord left,
            ComposeShareCleanupRecord right)
        {
            string leftFingerprint =
                left.Origin != null
                    ? left.Origin.AccountFingerprint
                    : string.Empty;
            string rightFingerprint =
                right.Origin != null
                    ? right.Origin.AccountFingerprint
                    : string.Empty;
            return string.Equals(
                       leftFingerprint,
                       rightFingerprint,
                       StringComparison.OrdinalIgnoreCase)
                   && string.Equals(
                       left.RelativeFolder,
                       right.RelativeFolder,
                       StringComparison.OrdinalIgnoreCase)
                   && string.Equals(
                       left.ShareId,
                       right.ShareId,
                       StringComparison.Ordinal);
        }
    }
}
