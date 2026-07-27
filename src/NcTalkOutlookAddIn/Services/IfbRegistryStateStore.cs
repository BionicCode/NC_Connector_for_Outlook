// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.Services
{
    internal sealed class IfbRegistryStateStore
    {
        private static readonly ProtectedJsonStateStoreDefinition<IfbRegistryState>
            StoreDefinition = new ProtectedJsonStateStoreDefinition<IfbRegistryState>(
                "ifb-registry-state",
                "NC4OL::IFB::RegistryState::v1",
                IsValid,
                () => new IfbRegistryState(),
                LogCategories.Ifb,
                new ProtectedJsonStateStoreMessages(
                    "Recovered the IFB registry state from its backup.",
                    "Failed to restore the IFB registry state backup.",
                    "IFB registry state recovery failed; the existing files were preserved.",
                    "IFB registry state is unreadable; existing data was preserved.",
                    "IFB registry state structure is invalid.",
                    "Failed to load IFB registry state file '"));

        private readonly ProtectedJsonStateStore<IfbRegistryState> _store;

        internal IfbRegistryStateStore(
            string dataDirectory,
            string profileScope)
        {
            _store = new ProtectedJsonStateStore<IfbRegistryState>(
                dataDirectory,
                profileScope,
                StoreDefinition);
        }

        internal IfbRegistryState Load()
        {
            return _store.Load();
        }

        internal void Save(IfbRegistryState state)
        {
            _store.Save(state);
        }

        private static bool IsValid(IfbRegistryState state)
        {
            if (state == null
                || state.Ownership == null
                || state.Ownership.Count > 2)
            {
                return false;
            }

            var keys = new HashSet<string>(
                StringComparer.OrdinalIgnoreCase);
            foreach (IfbRegistryOwnership ownership in state.Ownership)
            {
                if (ownership == null
                    || string.IsNullOrWhiteSpace(ownership.RegistryPath)
                    || string.IsNullOrWhiteSpace(ownership.ValueName)
                    || string.IsNullOrWhiteSpace(ownership.WrittenValue)
                    || !keys.Add(
                        ownership.RegistryPath
                        + "\n"
                        + ownership.ValueName))
                {
                    return false;
                }
            }
            return true;
        }
    }

    internal sealed class IfbRegistryState
    {
        public IfbRegistryState()
        {
            Ownership = new List<IfbRegistryOwnership>();
        }

        public List<IfbRegistryOwnership> Ownership { get; set; }
    }

    internal sealed class IfbRegistryOwnership
    {
        public string RegistryPath { get; set; }

        public string ValueName { get; set; }

        public bool OriginalExists { get; set; }

        public int OriginalKind { get; set; }

        public string OriginalValue { get; set; }

        public string WrittenValue { get; set; }
    }
}
