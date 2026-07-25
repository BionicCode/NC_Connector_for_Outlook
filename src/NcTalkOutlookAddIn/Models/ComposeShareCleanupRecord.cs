// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System.Runtime.InteropServices;

namespace NcTalkOutlookAddIn.Models
{
    [ComVisible(false)]
    public sealed class ComposeShareCleanupRecord
    {
        public string RelativeFolder { get; set; }

        public string ShareId { get; set; }

        public string ShareLabel { get; set; }

        public ComposeLifecycleOrigin Origin { get; set; }
    }
}
