// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using NcTalkOutlookAddIn.Services;

namespace NcTalkOutlookAddIn.Models
{
    internal sealed class TalkRoomTrackingSnapshot
    {
        internal TalkRoomTrackingSnapshot()
        {
            EntryId = string.Empty;
            StoreId = string.Empty;
            FolderId = string.Empty;
            GlobalAppointmentId = string.Empty;
            RoomToken = string.Empty;
        }

        internal string EntryId { get; set; }

        internal string StoreId { get; set; }

        internal string FolderId { get; set; }

        internal string GlobalAppointmentId { get; set; }

        internal string RoomToken { get; set; }

        internal bool IsEventConversation { get; set; }

        internal TalkServiceConfiguration Configuration { get; set; }
    }

    internal sealed class TalkRoomLifecycleRecord
    {
        public TalkRoomLifecycleRecord()
        {
            Id = Guid.NewGuid().ToString("N");
            EntryId = string.Empty;
            StoreId = string.Empty;
            FolderId = string.Empty;
            GlobalAppointmentId = string.Empty;
            RoomToken = string.Empty;
            ServerBaseUrl = string.Empty;
            AccountLogin = string.Empty;
            AccountId = string.Empty;
            CreatedUtc = DateTime.UtcNow;
            FirstMissingUtc = DateTime.MinValue;
            LastSeenUtc = DateTime.UtcNow;
            NextAttemptUtc = DateTime.MinValue;
        }

        public string Id { get; set; }

        public string EntryId { get; set; }

        public string StoreId { get; set; }

        public string FolderId { get; set; }

        public string GlobalAppointmentId { get; set; }

        public string RoomToken { get; set; }

        public bool IsEventConversation { get; set; }

        public string ServerBaseUrl { get; set; }

        public string AccountLogin { get; set; }

        public string AccountId { get; set; }

        public bool PendingDeletion { get; set; }

        public bool PolicyRequired { get; set; }

        public int AttemptCount { get; set; }

        public DateTime CreatedUtc { get; set; }

        public int MissingConfirmationCount { get; set; }

        public DateTime FirstMissingUtc { get; set; }

        public DateTime LastSeenUtc { get; set; }

        public DateTime NextAttemptUtc { get; set; }
    }

    internal sealed class TalkRoomReconcileResult
    {
        internal TalkRoomReconcileResult()
        {
            NextConfirmationUtc = DateTime.MinValue;
        }

        internal bool NeedsFollowUp { get; set; }

        internal DateTime NextConfirmationUtc { get; set; }

        internal int ConfirmedDeletionCount { get; set; }
    }
}
