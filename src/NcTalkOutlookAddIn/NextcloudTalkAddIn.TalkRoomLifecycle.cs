// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using NcTalkOutlookAddIn.Services;
using NcTalkOutlookAddIn.Settings;

namespace NcTalkOutlookAddIn
{
    // Persists requested Talk room deletions without enumerating Outlook calendars.
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

        internal bool QueueSavedTalkRoomDeletion(
            string roomToken,
            bool isEventConversation)
        {
            return QueueTalkRoomDeletion(
                roomToken,
                isEventConversation,
                true);
        }

        internal bool QueueUnsavedTalkRoomDeletion(
            string roomToken,
            bool isEventConversation)
        {
            return QueueTalkRoomDeletion(
                roomToken,
                isEventConversation,
                false);
        }

        private bool QueueTalkRoomDeletion(
            string roomToken,
            bool isEventConversation,
            bool policyRequired)
        {
            TalkRoomLifecycleCoordinator coordinator =
                _talkRoomLifecycleCoordinator;
            AddinSettings settings = _currentSettings != null
                ? _currentSettings.Clone()
                : null;
            if (coordinator == null || settings == null)
            {
                return false;
            }

            return coordinator.QueueDeletion(
                roomToken,
                isEventConversation,
                new TalkServiceConfiguration(
                    settings.ServerUrl,
                    settings.Username,
                    settings.AppPassword),
                policyRequired);
        }

        private void DisposeTalkRoomLifecycle()
        {
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
