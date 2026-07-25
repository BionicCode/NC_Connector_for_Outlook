// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using NcTalkOutlookAddIn.Utilities;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace NcTalkOutlookAddIn.Services
{
    // Keeps one Items event source per calendar folder instead of retaining every appointment RCW.
    internal sealed class TalkCalendarLifecycleMonitor : IDisposable
    {
        private readonly Action<Outlook.AppointmentItem> _onItemChanged;
        private readonly Action<string, string> _onItemRemoved;
        private readonly List<CalendarSubscription> _subscriptions =
            new List<CalendarSubscription>();
        private bool _setupComplete = true;
        private bool _disposed;

        internal TalkCalendarLifecycleMonitor(
            Outlook.Application application,
            Action<Outlook.AppointmentItem> onItemChanged,
            Action<string, string> onItemRemoved)
        {
            if (application == null)
            {
                throw new ArgumentNullException("application");
            }
            if (onItemChanged == null)
            {
                throw new ArgumentNullException("onItemChanged");
            }
            if (onItemRemoved == null)
            {
                throw new ArgumentNullException("onItemRemoved");
            }
            _onItemChanged = onItemChanged;
            _onItemRemoved = onItemRemoved;
            HookCalendars(application);
        }

        internal int FolderCount
        {
            get { return _subscriptions.Count; }
        }

        internal bool SetupComplete
        {
            get { return _setupComplete; }
        }

        internal bool TryScanAll(
            Action<Outlook.AppointmentItem> onAppointment)
        {
            if (_disposed
                || !_setupComplete
                || onAppointment == null
                || _subscriptions.Count == 0)
            {
                return false;
            }

            bool complete = true;
            for (int i = 0; i < _subscriptions.Count; i++)
            {
                CalendarSubscription subscription =
                    _subscriptions[i];
                if (subscription == null
                    || subscription.Items == null)
                {
                    complete = false;
                    continue;
                }

                int count;
                try
                {
                    count = subscription.Items.Count;
                }
                catch (Exception ex)
                {
                    complete = false;
                    DiagnosticsLogger.LogException(
                        LogCategories.Talk,
                        "A monitored calendar could not be counted during Talk lifecycle reconciliation.",
                        ex);
                    continue;
                }

                for (int itemIndex = 1;
                     itemIndex <= count;
                     itemIndex++)
                {
                    object item = null;
                    try
                    {
                        item = subscription.Items[itemIndex];
                        Outlook.AppointmentItem appointment =
                            item as Outlook.AppointmentItem;
                        if (appointment != null)
                        {
                            onAppointment(appointment);
                        }
                    }
                    catch (Exception ex)
                    {
                        complete = false;
                        DiagnosticsLogger.LogException(
                            LogCategories.Talk,
                            "A calendar item could not be inspected during Talk lifecycle reconciliation.",
                            ex);
                    }
                    finally
                    {
                        ComInteropScope.TryRelease(
                            item,
                            LogCategories.Talk,
                            "Failed to release a calendar item after Talk lifecycle reconciliation.");
                    }
                }
            }
            return complete;
        }

        private void HookCalendars(
            Outlook.Application application)
        {
            Outlook.NameSpace session = null;
            Outlook.Stores stores = null;
            try
            {
                session = application.Session;
                stores = session != null ? session.Stores : null;
                int count = stores != null ? stores.Count : 0;
                for (int i = 1; i <= count; i++)
                {
                    Outlook.Store store = null;
                    Outlook.MAPIFolder calendar = null;
                    try
                    {
                        store = stores[i];
                        calendar = store.GetDefaultFolder(
                            Outlook.OlDefaultFolders.olFolderCalendar);
                        HookCalendarTree(calendar);
                    }
                    catch (Exception ex)
                    {
                        _setupComplete = false;
                        DiagnosticsLogger.LogException(
                            LogCategories.Talk,
                            "A calendar store could not be monitored for Talk room lifecycle changes.",
                            ex);
                    }
                    finally
                    {
                        ComInteropScope.TryRelease(
                            calendar,
                            LogCategories.Talk,
                            "Failed to release a calendar folder after lifecycle monitoring setup.");
                        ComInteropScope.TryRelease(
                            store,
                            LogCategories.Talk,
                            "Failed to release a calendar store after lifecycle monitoring setup.");
                    }
                }
            }
            finally
            {
                ComInteropScope.TryRelease(
                    stores,
                    LogCategories.Talk,
                    "Failed to release calendar Stores after lifecycle monitoring setup.");
                ComInteropScope.TryRelease(
                    session,
                    LogCategories.Talk,
                    "Failed to release the Outlook session after lifecycle monitoring setup.");
            }
        }

        private void HookCalendarTree(
            Outlook.MAPIFolder folder)
        {
            if (folder == null)
            {
                return;
            }

            HookFolder(folder);
            Outlook.Folders children = null;
            try
            {
                children = folder.Folders;
                int count = children != null ? children.Count : 0;
                for (int i = 1; i <= count; i++)
                {
                    Outlook.MAPIFolder child = null;
                    try
                    {
                        child = children[i];
                        if (child != null
                            && child.DefaultItemType
                            == Outlook.OlItemType.olAppointmentItem)
                        {
                            HookCalendarTree(child);
                        }
                    }
                    catch (Exception ex)
                    {
                        _setupComplete = false;
                        DiagnosticsLogger.LogException(
                            LogCategories.Talk,
                            "A calendar subfolder could not be monitored.",
                            ex);
                    }
                    finally
                    {
                        ComInteropScope.TryRelease(
                            child,
                            LogCategories.Talk,
                            "Failed to release a calendar subfolder after lifecycle monitoring setup.");
                    }
                }
            }
            finally
            {
                ComInteropScope.TryRelease(
                    children,
                    LogCategories.Talk,
                    "Failed to release calendar subfolders after lifecycle monitoring setup.");
            }
        }

        private void HookFolder(
            Outlook.MAPIFolder folder)
        {
            Outlook.Items items = null;
            bool retained = false;
            try
            {
                string storeId = folder.StoreID ?? string.Empty;
                string folderId = folder.EntryID ?? string.Empty;
                items = folder.Items;
                if (items == null)
                {
                    return;
                }

                var subscription = new CalendarSubscription(
                    items,
                    storeId,
                    folderId);
                subscription.ItemAddHandler =
                    item => OnItemChanged(item);
                subscription.ItemChangeHandler =
                    item => OnItemChanged(item);
                subscription.ItemRemoveHandler =
                    () => OnItemRemoved(
                        subscription.StoreId,
                        subscription.FolderId);
                items.ItemAdd += subscription.ItemAddHandler;
                items.ItemChange += subscription.ItemChangeHandler;
                items.ItemRemove += subscription.ItemRemoveHandler;
                _subscriptions.Add(subscription);
                retained = true;
            }
            finally
            {
                if (!retained)
                {
                    ComInteropScope.TryRelease(
                        items,
                        LogCategories.Talk,
                        "Failed to release an unmonitored calendar Items collection.");
                }
            }
        }

        private void OnItemChanged(object item)
        {
            if (_disposed)
            {
                return;
            }
            Outlook.AppointmentItem appointment =
                item as Outlook.AppointmentItem;
            if (appointment != null)
            {
                _onItemChanged(appointment);
            }
        }

        private void OnItemRemoved(
            string storeId,
            string folderId)
        {
            if (!_disposed)
            {
                _onItemRemoved(storeId, folderId);
            }
        }

        public void Dispose()
        {
            _disposed = true;
            for (int i = 0; i < _subscriptions.Count; i++)
            {
                CalendarSubscription subscription =
                    _subscriptions[i];
                if (subscription == null
                    || subscription.Items == null)
                {
                    continue;
                }
                try
                {
                    subscription.Items.ItemAdd -=
                        subscription.ItemAddHandler;
                    subscription.Items.ItemChange -=
                        subscription.ItemChangeHandler;
                    subscription.Items.ItemRemove -=
                        subscription.ItemRemoveHandler;
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(
                        LogCategories.Talk,
                        "Calendar lifecycle event handlers could not be removed.",
                        ex);
                }
                ComInteropScope.TryRelease(
                    subscription.Items,
                    LogCategories.Talk,
                    "Failed to release a monitored calendar Items collection.");
            }
            _subscriptions.Clear();
        }

        private sealed class CalendarSubscription
        {
            internal CalendarSubscription(
                Outlook.Items items,
                string storeId,
                string folderId)
            {
                Items = items;
                StoreId = storeId ?? string.Empty;
                FolderId = folderId ?? string.Empty;
            }

            internal Outlook.Items Items { get; private set; }

            internal string StoreId { get; private set; }

            internal string FolderId { get; private set; }

            internal Outlook.ItemsEvents_ItemAddEventHandler
                ItemAddHandler { get; set; }

            internal Outlook.ItemsEvents_ItemChangeEventHandler
                ItemChangeHandler { get; set; }

            internal Outlook.ItemsEvents_ItemRemoveEventHandler
                ItemRemoveHandler { get; set; }
        }
    }
}
