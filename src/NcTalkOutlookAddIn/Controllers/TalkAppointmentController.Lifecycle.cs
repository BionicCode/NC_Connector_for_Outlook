// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Services;
using NcTalkOutlookAddIn.Utilities;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace NcTalkOutlookAddIn.Controllers
{
    // Captures stable appointment identity before lifecycle work leaves the Outlook thread.
    internal sealed partial class TalkAppointmentController
    {
        internal TalkRoomTrackingSnapshot CaptureTrackingSnapshot(
            Outlook.AppointmentItem appointment,
            string roomToken,
            bool isEventConversation)
        {
            if (appointment == null
                || _owner.CurrentSettings == null)
            {
                return null;
            }

            var configuration = new TalkServiceConfiguration(
                _owner.CurrentSettings.ServerUrl,
                _owner.CurrentSettings.Username,
                _owner.CurrentSettings.AppPassword);
            if (!configuration.IsComplete())
            {
                return null;
            }

            string storeId;
            string folderId;
            ReadFolderIdentity(
                appointment,
                out storeId,
                out folderId);
            return new TalkRoomTrackingSnapshot
            {
                EntryId = ReadEntryId(appointment),
                StoreId = storeId,
                FolderId = folderId,
                GlobalAppointmentId =
                    ReadGlobalAppointmentId(appointment),
                RoomToken = (roomToken ?? string.Empty).Trim(),
                IsEventConversation = isEventConversation,
                Configuration = configuration
            };
        }

        private static string ReadEntryId(
            Outlook.AppointmentItem appointment)
        {
            try
            {
                return appointment.EntryID ?? string.Empty;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(
                    LogCategories.Talk,
                    "Failed to read the appointment EntryID for room lifecycle tracking.",
                    ex);
                return string.Empty;
            }
        }

        private static string ReadGlobalAppointmentId(
            Outlook.AppointmentItem appointment)
        {
            try
            {
                return appointment.GlobalAppointmentID ?? string.Empty;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(
                    LogCategories.Talk,
                    "Failed to read the appointment global ID for room lifecycle tracking.",
                    ex);
                return string.Empty;
            }
        }

        private static void ReadFolderIdentity(
            Outlook.AppointmentItem appointment,
            out string storeId,
            out string folderId)
        {
            storeId = string.Empty;
            folderId = string.Empty;
            Outlook.MAPIFolder folder = null;
            try
            {
                folder = appointment.Parent as Outlook.MAPIFolder;
                if (folder != null)
                {
                    storeId = folder.StoreID ?? string.Empty;
                    folderId = folder.EntryID ?? string.Empty;
                }
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(
                    LogCategories.Talk,
                    "Failed to read the appointment folder identity for room lifecycle tracking.",
                    ex);
            }
            finally
            {
                if (folder != null)
                {
                    ComInteropScope.TryRelease(
                        folder,
                        LogCategories.Talk,
                        "Failed to release the appointment folder after lifecycle capture.");
                }
            }
        }
    }
}
