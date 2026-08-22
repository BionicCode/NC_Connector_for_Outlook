// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System.Collections.Generic;
using System.Globalization;
using NcTalkOutlookAddIn.Controllers;
using NcTalkOutlookAddIn.Models;

namespace NcTalkOutlookAddIn
{
    public sealed partial class NextcloudTalkAddIn
    {
        internal sealed partial class MailComposeSubscription
        {
            private readonly ComposeShareCleanupTracker
                _shareCleanupTracker =
                    new ComposeShareCleanupTracker();

            internal void ArmShareCleanup(
                FileLinkResult result,
                ComposeLifecycleOrigin origin)
            {
                if (_disposed || result == null)
                {
                    return;
                }

                string relativeFolder =
                    string.IsNullOrWhiteSpace(result.RelativePath)
                        ? string.Empty
                        : result.RelativePath.Trim();
                var record = new ComposeShareCleanupRecord
                {
                    RelativeFolder = relativeFolder,
                    ShareId = result.ShareId ?? string.Empty,
                    ShareLabel = result.FolderName ?? string.Empty,
                    Origin = origin != null ? origin.Clone() : null
                };
                if (!_shareCleanupTracker.Arm(record))
                {
                    LogFileLink(
                        "Compose share cleanup arm skipped (composeKey="
                        + _composeKey
                        + ", reason="
                        + (string.IsNullOrWhiteSpace(relativeFolder)
                            ? "missing_relative_folder"
                            : "already_tracked")
                        + ").");
                    return;
                }

                LogFileLink(
                    "Compose share cleanup armed until Outlook writes the item (composeKey="
                    + _composeKey
                    + ", armedCount="
                    + _shareCleanupTracker.Count.ToString(
                        CultureInfo.InvariantCulture)
                    + ").");
            }

            private void OnAfterWrite()
            {
                if (_disposed)
                {
                    return;
                }

                // AfterWrite covers Save, AutoSave, and Send. The share now belongs to a
                // persisted message and must not be removed after a later discard.
                int released = _shareCleanupTracker.ReleaseAll();
                if (released == 0)
                {
                    return;
                }

                LogFileLink(
                    "Compose share cleanup released after Outlook write (composeKey="
                    + _composeKey
                    + ", released="
                    + released.ToString(CultureInfo.InvariantCulture)
                    + ").");
            }

            private void OnUnload()
            {
                CompleteComposeShareCleanup(
                    "Compose item unloaded",
                    "compose_unloaded_without_write",
                    false);
            }

            private void OnInspectorClosed()
            {
                CompleteComposeShareCleanup(
                    "Compose Inspector closed",
                    "compose_inspector_closed_without_write",
                    true);
            }

            private void CompleteComposeShareCleanup(
                string eventLabel,
                string cleanupReason,
                bool detachItemEvents)
            {
                if (_disposed)
                {
                    return;
                }

                List<ComposeShareCleanupRecord> pending =
                    _shareCleanupTracker.Drain();
                LogFileLink(
                    (eventLabel ?? "Compose lifecycle completed")
                    + " (composeKey="
                    + _composeKey
                    + ", cleanupCount="
                    + pending.Count.ToString(CultureInfo.InvariantCulture)
                    + ").");

                Dispose(detachItemEvents);
                if (pending.Count > 0)
                {
                    _owner.QueueCreatedShareCleanup(
                        _composeKey,
                        pending,
                        cleanupReason);
                }
            }
        }
    }
}
