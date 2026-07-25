// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System.Collections.Generic;
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

        private void InitializeComposeLifecycle(string outlookProfileName)
        {
            DisposeComposeLifecycle();
            _composeShareLifecycleController.AttachOwner(this);
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
            _composeShareLifecycleController.TryDeleteComposeShareFolder(
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
                    Origin = origin != null ? origin.Clone() : null
                },
                reason);
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

        internal bool IsMailComposeSurfaceOpen(Outlook.MailItem mail)
        {
            return _mailInteropController.IsActiveInlineResponse(mail)
                   || _mailInteropController.IsMailOpenInAnyInspector(mail);
        }
    }
}
