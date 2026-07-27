// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.Services
{
    internal sealed class TalkRoomLifecycleStore
    {
        private static readonly ProtectedJsonStateStoreDefinition<TalkRoomLifecycleState>
            StoreDefinition = new ProtectedJsonStateStoreDefinition<TalkRoomLifecycleState>(
                "talk-room-lifecycle",
                "NC4OL::TalkRoomLifecycle::v1",
                IsValid,
                () => new TalkRoomLifecycleState(),
                LogCategories.Talk,
                new ProtectedJsonStateStoreMessages(
                    "Recovered the Talk room deletion queue from its backup.",
                    "Failed to restore the Talk room deletion queue from its backup.",
                    "Talk room deletion queue recovery failed; writes are disabled to preserve the existing files.",
                    "Talk room deletion queue is unreadable; existing data was preserved.",
                    "Talk room deletion queue structure is invalid.",
                    "Failed to load Talk room deletion queue file '"));

        private readonly ProtectedJsonStateStore<TalkRoomLifecycleState> _store;

        internal TalkRoomLifecycleStore(
            string dataDirectory,
            string profileScope)
        {
            _store = new ProtectedJsonStateStore<TalkRoomLifecycleState>(
                dataDirectory,
                profileScope,
                StoreDefinition);
        }

        internal TalkRoomLifecycleState Load()
        {
            return _store.Load();
        }

        internal void Save(TalkRoomLifecycleState state)
        {
            _store.Save(state);
        }

        private static bool IsValid(TalkRoomLifecycleState state)
        {
            if (state == null || state.Records == null)
            {
                return false;
            }
            var ids = new HashSet<string>(StringComparer.Ordinal);
            for (int i = 0; i < state.Records.Count; i++)
            {
                TalkRoomLifecycleRecord record = state.Records[i];
                if (record == null
                    || string.IsNullOrWhiteSpace(record.Id)
                    || string.IsNullOrWhiteSpace(record.RoomToken)
                    || string.IsNullOrWhiteSpace(record.ServerBaseUrl)
                    || !ids.Add(record.Id))
                {
                    return false;
                }
            }
            return true;
        }
    }

    internal sealed class TalkRoomLifecycleState
    {
        public TalkRoomLifecycleState()
        {
            Records = new List<TalkRoomLifecycleRecord>();
        }

        public List<TalkRoomLifecycleRecord> Records { get; set; }
    }
}
