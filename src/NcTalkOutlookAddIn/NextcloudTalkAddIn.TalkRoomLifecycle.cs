// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Threading.Tasks;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Services;
using NcTalkOutlookAddIn.Settings;
using NcTalkOutlookAddIn.Utilities;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace NcTalkOutlookAddIn
{
    // Connects Outlook appointments to the compact persistent room index.
    public sealed partial class NextcloudTalkAddIn
    {
        private TalkRoomLifecycleCoordinator
            _talkRoomLifecycleCoordinator;

        private void InitializeTalkRoomLifecycle(
            string dataDirectory,
            string outlookProfileScope)
        {
            _talkRoomLifecycleCoordinator =
                new TalkRoomLifecycleCoordinator(
                    dataDirectory,
                    outlookProfileScope,
                    () => _currentSettings != null
                        ? _currentSettings.Clone()
                        : new AddinSettings(),
                    ShouldDeleteTalkRoomOnSavedEventDelete);
            _talkRoomLifecycleCoordinator.StartPendingProcessing();
        }

        internal void TrackTalkRoomAppointment(
            Outlook.AppointmentItem appointment,
            string roomToken,
            bool isEventConversation)
        {
            TalkRoomLifecycleCoordinator coordinator =
                _talkRoomLifecycleCoordinator;
            if (coordinator == null)
            {
                return;
            }

            TalkRoomTrackingSnapshot snapshot =
                _talkAppointmentController.CaptureTrackingSnapshot(
                    appointment,
                    roomToken,
                    isEventConversation);
            if (snapshot != null
                && !string.IsNullOrWhiteSpace(snapshot.EntryId))
            {
                coordinator.Track(snapshot);
            }
        }

        internal void QueueUnsavedTalkRoomDeletion(
            string roomToken,
            bool isEventConversation)
        {
            TalkRoomLifecycleCoordinator coordinator =
                _talkRoomLifecycleCoordinator;
            AddinSettings settings = _currentSettings != null
                ? _currentSettings.Clone()
                : null;
            if (coordinator == null || settings == null)
            {
                return;
            }

            coordinator.QueueUnconditionalDeletion(
                roomToken,
                isEventConversation,
                new TalkServiceConfiguration(
                    settings.ServerUrl,
                    settings.Username,
                    settings.AppPassword));
        }

        private void StartTalkRoomLifecycleRecovery()
        {
            Task recovery = RunOnOutlookUiThreadAsync(
                () => EnsureTalkCalendarLifecycleMonitor(true));
            recovery.ContinueWith(
                task => DiagnosticsLogger.LogException(
                    LogCategories.Talk,
                    "Talk room lifecycle recovery failed.",
                    task.Exception),
                TaskContinuationOptions.OnlyOnFaulted);
        }

        private void DisposeTalkRoomLifecycle()
        {
            DisposeTalkCalendarLifecycleMonitor();
            TalkRoomLifecycleCoordinator coordinator =
                _talkRoomLifecycleCoordinator;
            _talkRoomLifecycleCoordinator = null;
            if (coordinator != null)
            {
                coordinator.Dispose();
            }
        }
    }
}
