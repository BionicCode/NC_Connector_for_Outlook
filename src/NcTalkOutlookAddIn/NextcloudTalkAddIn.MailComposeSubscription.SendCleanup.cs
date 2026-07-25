// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Windows.Forms;
using NcTalkOutlookAddIn.Controllers;
using NcTalkOutlookAddIn.Services;
using NcTalkOutlookAddIn.Utilities;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace NcTalkOutlookAddIn
{
    public sealed partial class NextcloudTalkAddIn
    {
        internal sealed partial class MailComposeSubscription
        {
            private const int SurfaceCloseVerificationMaxAttempts = 8;

            private void OnSend(ref bool cancel)
            {
                if (_disposed || cancel)
                {
                    return;
                }
                if (!TryValidateAttachmentPolicyBeforeSend(ref cancel)
                    || !TryFinalizeEmailSignatureBeforeSend(ref cancel))
                {
                    return;
                }

                CapturePasswordDispatchRecipients();
                CapturePasswordDispatchSender();

                if (_passwordDispatchQueue.Count > 0)
                {
                    if (!_owner.TryArmPendingPasswordDrafts(
                        _mail,
                        _composeKey,
                        _passwordDispatchQueue))
                    {
                        cancel = true;
                        MessageBox.Show(
                            Strings.SharingPasswordMailPrepareFailed,
                            Strings.DialogTitle,
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning);
                        return;
                    }
                }

                _cleanupGraceTimer.Stop();
                LogFileLink(
                    "Compose send armed for positive Sent confirmation (composeKey="
                    + _composeKey
                    + ", passwordQueued="
                    + _passwordDispatchQueue.Count.ToString(CultureInfo.InvariantCulture)
                    + ").");
            }

            private void OnClose(ref bool cancel)
            {
                if (_disposed)
                {
                    return;
                }

                ScheduleSurfaceCloseVerification(
                    cancel ? "close_already_cancelled" : "close_event");
            }

            private void ScheduleSurfaceCloseVerification(string reason)
            {
                if (_disposed)
                {
                    return;
                }

                _surfaceCloseVerificationAttempts = 0;
                _cleanupGraceTimer.Stop();
                _cleanupGraceTimer.Interval = 250;
                _cleanupGraceTimer.Start();
                LogFileLink(
                    "Compose surface-close verification scheduled (composeKey="
                    + _composeKey
                    + ", reason="
                    + (reason ?? string.Empty)
                    + ").");
            }

            private void OnCleanupGraceTimerTick(object sender, EventArgs e)
            {
                _cleanupGraceTimer.Stop();
                if (_disposed)
                {
                    return;
                }

                _surfaceCloseVerificationAttempts++;
                if (_owner.IsMailComposeSurfaceOpen(_mail))
                {
                    _surfaceCloseVerificationAttempts = 0;
                    LogFileLink(
                        "Compose surface remains open; close cleanup cancelled (composeKey="
                        + _composeKey
                        + ").");
                    return;
                }

                if (_surfaceCloseVerificationAttempts
                    < SurfaceCloseVerificationMaxAttempts)
                {
                    _cleanupGraceTimer.Start();
                    return;
                }

                LogFileLink(
                    "Compose surface closed; unknown state retained without destructive cleanup (composeKey="
                    + _composeKey
                    + ").");
                Dispose();
            }

            private void CapturePasswordDispatchRecipients()
            {
                if (_passwordDispatchQueue.Count == 0)
                {
                    return;
                }
                string to;
                string cc;
                string bcc;
                bool capturedFromRecipients =
                    TryCaptureRecipientListsFromRecipientsCollection(
                        out to,
                        out cc,
                        out bcc);
                if (!capturedFromRecipients)
                {
                    to = ComposeShareLifecycleController.BuildNormalizedRecipientCsv(
                        ReadMailRecipientList("To"));
                    cc = ComposeShareLifecycleController.BuildNormalizedRecipientCsv(
                        ReadMailRecipientList("CC"));
                    bcc = ComposeShareLifecycleController.BuildNormalizedRecipientCsv(
                        ReadMailRecipientList("BCC"));
                }
                for (int i = 0; i < _passwordDispatchQueue.Count; i++)
                {
                    _passwordDispatchQueue[i].To = to;
                    _passwordDispatchQueue[i].Cc = cc;
                    _passwordDispatchQueue[i].Bcc = bcc;
                }

                LogFileLink(
                    "Separate password recipients captured (composeKey="
                    + _composeKey
                    + ", queued="
                    + _passwordDispatchQueue.Count.ToString(CultureInfo.InvariantCulture)
                    + ", to="
                    + CountRecipients(to).ToString(CultureInfo.InvariantCulture)
                    + ", cc="
                    + CountRecipients(cc).ToString(CultureInfo.InvariantCulture)
                    + ", bcc="
                    + CountRecipients(bcc).ToString(CultureInfo.InvariantCulture)
                    + ", source="
                    + (capturedFromRecipients
                        ? "recipients_collection"
                        : "mail_fields")
                    + ").");
            }

            private void CapturePasswordDispatchSender()
            {
                if (_passwordDispatchQueue.Count == 0)
                {
                    return;
                }

                string senderEmail =
                    EmailSignaturePolicyService.NormalizeEmail(
                        ResolveCurrentSenderEmail());
                string accountSmtp =
                    EmailSignaturePolicyService.NormalizeEmail(
                        OutlookRecipientResolverController
                            .ResolveSendUsingAccountSmtpAddress(
                                _mail,
                                LogCategories.Core,
                                "compose",
                                string.Empty));
                string sentOnBehalfOfName =
                    ReadCurrentSentOnBehalfOfName();
                for (int i = 0; i < _passwordDispatchQueue.Count; i++)
                {
                    _passwordDispatchQueue[i].SenderEmail = senderEmail;
                    _passwordDispatchQueue[i]
                        .SendUsingAccountSmtpAddress = accountSmtp;
                    _passwordDispatchQueue[i]
                        .SentOnBehalfOfName = sentOnBehalfOfName;
                }

                LogFileLink(
                    "Separate password sender captured (composeKey="
                    + _composeKey
                    + ", queued="
                    + _passwordDispatchQueue.Count.ToString(CultureInfo.InvariantCulture)
                    + ", hasSender="
                    + (!string.IsNullOrWhiteSpace(senderEmail))
                        .ToString(CultureInfo.InvariantCulture)
                    + ", hasAccount="
                    + (!string.IsNullOrWhiteSpace(accountSmtp))
                        .ToString(CultureInfo.InvariantCulture)
                    + ", sentOnBehalf="
                    + (!string.IsNullOrWhiteSpace(sentOnBehalfOfName))
                        .ToString(CultureInfo.InvariantCulture)
                    + ").");
            }

            private string ReadCurrentSentOnBehalfOfName()
            {
                try
                {
                    return _mail != null
                        ? (_mail.SentOnBehalfOfName ?? string.Empty).Trim()
                        : string.Empty;
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(
                        LogCategories.FileLink,
                        "Failed to read compose sent-on-behalf name for separate password mail (composeKey=" + _composeKey + ").",
                        ex);
                    return string.Empty;
                }
            }

            private bool TryCaptureRecipientListsFromRecipientsCollection(
                out string to,
                out string cc,
                out string bcc)
            {
                to = string.Empty;
                cc = string.Empty;
                bcc = string.Empty;
                if (_mail == null)
                {
                    return false;
                }

                var toRecipients = new List<string>();
                var ccRecipients = new List<string>();
                var bccRecipients = new List<string>();
                Outlook.Recipients recipients = null;
                try
                {
                    recipients = _mail.Recipients;
                    if (recipients == null)
                    {
                        return false;
                    }

                    int count;
                    try
                    {
                        count = recipients.Count;
                    }
                    catch (Exception ex)
                    {
                        DiagnosticsLogger.LogException(
                            LogCategories.FileLink,
                            "Failed to read compose Recipients.Count (composeKey=" + _composeKey + ").",
                            ex);
                        count = 0;
                    }

                    for (int i = 1; i <= count; i++)
                    {
                        Outlook.Recipient recipient = null;
                        try
                        {
                            recipient = recipients[i];
                            if (recipient == null)
                            {
                                continue;
                            }

                            string address =
                                TryGetRecipientSmtpAddress(recipient);
                            if (string.IsNullOrWhiteSpace(address))
                            {
                                try
                                {
                                    address =
                                        ComposeShareLifecycleController
                                            .NormalizeRecipientAddress(
                                                recipient.Address);
                                }
                                catch (Exception ex)
                                {
                                    DiagnosticsLogger.LogException(
                                        LogCategories.FileLink,
                                        "Failed to read compose recipient.Address (composeKey=" + _composeKey + ").",
                                        ex);
                                    address = string.Empty;
                                }
                            }
                            if (string.IsNullOrWhiteSpace(address))
                            {
                                continue;
                            }

                            int recipientType;
                            try
                            {
                                recipientType = recipient.Type;
                            }
                            catch (Exception ex)
                            {
                                DiagnosticsLogger.LogException(
                                    LogCategories.FileLink,
                                    "Failed to read compose recipient.Type (composeKey=" + _composeKey + ").",
                                    ex);
                                recipientType =
                                    (int)Outlook.OlMailRecipientType.olTo;
                            }

                            if (recipientType
                                == (int)Outlook.OlMailRecipientType.olCC)
                            {
                                ComposeShareLifecycleController
                                    .AddUniqueRecipient(
                                        ccRecipients,
                                        address);
                            }
                            else if (recipientType
                                     == (int)Outlook.OlMailRecipientType
                                         .olBCC)
                            {
                                ComposeShareLifecycleController
                                    .AddUniqueRecipient(
                                        bccRecipients,
                                        address);
                            }
                            else
                            {
                                ComposeShareLifecycleController
                                    .AddUniqueRecipient(
                                        toRecipients,
                                        address);
                            }
                        }
                        finally
                        {
                            ComInteropScope.TryRelease(
                                recipient,
                                LogCategories.FileLink,
                                "Failed to release compose Recipient COM object.");
                        }
                    }
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(
                        LogCategories.FileLink,
                        "Failed to capture compose recipients from Recipients collection (composeKey=" + _composeKey + ").",
                        ex);
                    return false;
                }
                finally
                {
                    ComInteropScope.TryRelease(
                        recipients,
                        LogCategories.FileLink,
                        "Failed to release compose Recipients COM object.");
                }

                to = toRecipients.Count == 0
                    ? string.Empty
                    : string.Join("; ", toRecipients.ToArray());
                cc = ccRecipients.Count == 0
                    ? string.Empty
                    : string.Join("; ", ccRecipients.ToArray());
                bcc = bccRecipients.Count == 0
                    ? string.Empty
                    : string.Join("; ", bccRecipients.ToArray());
                return toRecipients.Count
                       + ccRecipients.Count
                       + bccRecipients.Count > 0;
            }

            private string ReadMailRecipientList(string fieldName)
            {
                try
                {
                    if (_mail == null)
                    {
                        return string.Empty;
                    }

                    switch ((fieldName ?? string.Empty)
                        .Trim()
                        .ToUpperInvariant())
                    {
                        case "TO":
                            return _mail.To ?? string.Empty;
                        case "CC":
                            return _mail.CC ?? string.Empty;
                        case "BCC":
                            return _mail.BCC ?? string.Empty;
                        default:
                            return string.Empty;
                    }
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(
                        LogCategories.FileLink,
                        "Failed to read compose recipient field '" + (fieldName ?? string.Empty) + "' (composeKey=" + _composeKey + ").",
                        ex);
                    return string.Empty;
                }
            }

            private static int CountRecipients(string csv)
            {
                return ComposeShareLifecycleController
                    .CountRecipientsInCsv(csv);
            }
        }
    }
}
