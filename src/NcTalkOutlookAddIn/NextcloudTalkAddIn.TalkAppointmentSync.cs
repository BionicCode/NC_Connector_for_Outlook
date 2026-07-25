// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Threading.Tasks;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Services;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn
{
    // Owns the background queue for appointment-to-Talk synchronization.
    public sealed partial class NextcloudTalkAddIn
    {
        private TalkAppointmentSyncCoordinator _talkAppointmentSyncCoordinator;
        private string _outlookProfileScope = string.Empty;

        internal string OutlookProfileScope
        {
            get
            {
                return string.IsNullOrWhiteSpace(_outlookProfileScope)
                    ? "default"
                    : _outlookProfileScope;
            }
        }

        private void InitializeTalkAppointmentSync(string outlookProfileScope)
        {
            _outlookProfileScope = string.IsNullOrWhiteSpace(outlookProfileScope)
                ? "default"
                : outlookProfileScope.Trim();
            _talkAppointmentSyncCoordinator = new TalkAppointmentSyncCoordinator(
                _talkAppointmentController.ExecuteRemoteSync,
                DispatchTalkAppointmentSyncResult);
        }

        internal void QueueTalkAppointmentSync(
            TalkAppointmentSyncSnapshot snapshot)
        {
            TalkAppointmentSyncCoordinator coordinator =
                _talkAppointmentSyncCoordinator;
            if (coordinator == null || snapshot == null)
            {
                return;
            }
            coordinator.Queue(snapshot);
        }

        private async Task DispatchTalkAppointmentSyncResult(
            TalkAppointmentSyncResult result)
        {
            bool handoffPersisted = await RunOnOutlookUiThreadAsync(
                () => ApplyTalkAppointmentSyncResult(result));
            if (!handoffPersisted
                || result == null
                || result.HandoffConfiguration == null)
            {
                return;
            }

            await Task.Run(
                () =>
                {
                    try
                    {
                        bool left = new TalkService(
                            result.HandoffConfiguration).LeaveRoom(
                                result.RoomToken);
                        LogTalkMessage(
                            "Moderator handoff completed (leftSelf="
                            + left
                            + ").");
                    }
                    catch (Exception ex)
                    {
                        DiagnosticsLogger.LogException(
                            LogCategories.Talk,
                            "Moderator was promoted and persisted, but leaving the room failed.",
                            ex);
                    }
                });
        }

        private bool ApplyTalkAppointmentSyncResult(
            TalkAppointmentSyncResult result)
        {
            if (result == null || string.IsNullOrWhiteSpace(result.RoomToken))
            {
                return false;
            }

            AppointmentSubscription subscription;
            if (_subscriptionByToken.TryGetValue(
                    result.RoomToken,
                    out subscription)
                && subscription != null)
            {
                return subscription.ApplyRemoteSyncResult(result);
            }
            return false;
        }

        private void DisposeTalkAppointmentSync()
        {
            TalkAppointmentSyncCoordinator coordinator =
                _talkAppointmentSyncCoordinator;
            _talkAppointmentSyncCoordinator = null;
            if (coordinator != null)
            {
                coordinator.Dispose();
            }
        }
    }
}
