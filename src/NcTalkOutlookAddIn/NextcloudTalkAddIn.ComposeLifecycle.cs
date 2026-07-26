// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System.Collections.Generic;
using System.Threading.Tasks;
using NcTalkOutlookAddIn.Controllers;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Settings;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace NcTalkOutlookAddIn
{
    public sealed partial class NextcloudTalkAddIn
    {
        private PendingPasswordDraftController
            _pendingPasswordDraftController;

        private void InitializeComposeLifecycle()
        {
            DisposeComposeLifecycle();
            _pendingPasswordDraftController =
                new PendingPasswordDraftController(
                    this,
                    _composeShareLifecycleController);
            _pendingPasswordDraftController.Start();
        }

        private void DisposeComposeLifecycle()
        {
            PendingPasswordDraftController controller =
                _pendingPasswordDraftController;
            _pendingPasswordDraftController = null;
            if (controller != null)
            {
                controller.Dispose();
            }
        }

        internal bool TryArmPendingPasswordDrafts(
            Outlook.MailItem primary,
            string composeKey,
            List<SeparatePasswordDispatchEntry> queue)
        {
            if (_pendingPasswordDraftController == null)
            {
                return false;
            }
            return _pendingPasswordDraftController.TryArm(
                primary,
                composeKey,
                queue);
        }

        internal void QueueCreatedShareCleanup(
            string composeKey,
            FileLinkResult result,
            ComposeLifecycleOrigin origin,
            string reason)
        {
            QueueCreatedShareCleanup(
                composeKey,
                new List<ComposeShareCleanupRecord>
                {
                    new ComposeShareCleanupRecord
                    {
                        RelativeFolder = result != null
                            ? result.RelativePath
                            : string.Empty,
                        ShareId = result != null
                            ? result.ShareId
                            : string.Empty,
                        ShareLabel = result != null
                            ? result.FolderName
                            : string.Empty,
                        Origin = origin
                    }
                },
                reason);
        }

        internal void QueueCreatedShareCleanup(
            string composeKey,
            List<ComposeShareCleanupRecord> records,
            string reason)
        {
            var pending = new List<ComposeShareCleanupRecord>();
            if (records != null)
            {
                for (int i = 0; i < records.Count; i++)
                {
                    ComposeShareCleanupRecord record = records[i];
                    if (record == null
                        || string.IsNullOrWhiteSpace(
                            record.RelativeFolder))
                    {
                        continue;
                    }
                    pending.Add(
                        new ComposeShareCleanupRecord
                        {
                            RelativeFolder =
                                record.RelativeFolder.Trim(),
                            ShareId = record.ShareId
                                ?? string.Empty,
                            ShareLabel = record.ShareLabel
                                ?? string.Empty,
                            Origin = record.Origin != null
                                ? record.Origin.Clone()
                                : null
                        });
                }
            }
            if (pending.Count == 0)
            {
                return;
            }

            LogFileLinkMessage(
                "Compose share cleanup queued (composeKey="
                + (composeKey ?? string.Empty)
                + ", count="
                + pending.Count
                + ", reason="
                + (reason ?? string.Empty)
                + ").");
            Task.Run(
                () =>
                {
                    for (int i = 0; i < pending.Count; i++)
                    {
                        _composeShareLifecycleController
                            .TryDeleteComposeShareFolder(
                                pending[i],
                                reason);
                    }
                });
        }

        internal void CaptureSeparatePasswordSignatureSnapshot(
            List<SeparatePasswordDispatchEntry> queue,
            BackendPolicyStatus policyStatus,
            AddinSettings settings,
            string composeKey)
        {
            _composeShareLifecycleController
                .CaptureSeparatePasswordSignatureSnapshot(
                    queue,
                    policyStatus,
                    settings,
                    composeKey);
        }

        internal bool TryInsertHtmlIntoMail(
            Outlook.MailItem mail,
            string html)
        {
            return _mailInteropController.InsertHtmlIntoMail(mail, html);
        }

        internal bool TryInsertPlainTextIntoMail(
            Outlook.MailItem mail,
            string plainText)
        {
            return _mailInteropController.InsertPlainTextIntoMail(
                mail,
                plainText);
        }

    }
}
