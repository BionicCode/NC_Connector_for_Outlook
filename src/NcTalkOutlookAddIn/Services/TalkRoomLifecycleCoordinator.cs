// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Threading;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Settings;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.Services
{
    // Owns the compact active-room index, absence confirmation and durable retries.
    internal sealed class TalkRoomLifecycleCoordinator : IDisposable
    {
        private static readonly TimeSpan MissingConfirmationDelay =
            TimeSpan.FromSeconds(30);

        private readonly object _syncRoot = new object();
        private readonly TalkRoomLifecycleStore _store;
        private readonly TalkRoomLifecycleState _state;
        private readonly Func<AddinSettings> _getCurrentSettings;
        private readonly Func<bool> _isSavedEventDeletionEnabled;
        private readonly Timer _timer;
        private int _workerRunning;
        private bool _disposed;

        internal TalkRoomLifecycleCoordinator(
            string dataDirectory,
            string profileScope,
            Func<AddinSettings> getCurrentSettings,
            Func<bool> isSavedEventDeletionEnabled)
        {
            if (getCurrentSettings == null)
            {
                throw new ArgumentNullException("getCurrentSettings");
            }
            if (isSavedEventDeletionEnabled == null)
            {
                throw new ArgumentNullException("isSavedEventDeletionEnabled");
            }

            _store = new TalkRoomLifecycleStore(dataDirectory, profileScope);
            _state = _store.Load() ?? new TalkRoomLifecycleState();
            if (_state.Records == null)
            {
                _state.Records = new List<TalkRoomLifecycleRecord>();
            }
            if (_state.SafetyVersion < 2)
            {
                for (int i = 0; i < _state.Records.Count; i++)
                {
                    TalkRoomLifecycleRecord record = _state.Records[i];
                    if (record != null
                        && record.PendingDeletion
                        && record.PolicyRequired)
                    {
                        ResetMissing(record);
                    }
                }
                _state.SafetyVersion = 2;
                _store.Save(_state);
            }
            _getCurrentSettings = getCurrentSettings;
            _isSavedEventDeletionEnabled = isSavedEventDeletionEnabled;
            _timer = new Timer(
                OnTimer,
                null,
                Timeout.Infinite,
                Timeout.Infinite);
        }

        internal void StartPendingProcessing()
        {
            ScheduleNext();
        }

        internal void Track(TalkRoomTrackingSnapshot snapshot)
        {
            if (!IsUsable(snapshot))
            {
                return;
            }

            try
            {
                string baseUrl = NormalizeBaseUrl(snapshot.Configuration);
                string login = NormalizeLogin(snapshot.Configuration);
                lock (_syncRoot)
                {
                    TalkRoomLifecycleRecord record =
                        FindSnapshotLocked(
                            snapshot,
                            baseUrl,
                            login,
                            string.Empty);
                    if (record == null)
                    {
                        record = NewRecord(baseUrl, login, string.Empty);
                    }
                    ApplySnapshot(record, snapshot);
                    ResetMissing(record);
                    SaveLocked();
                }
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(
                    LogCategories.Talk,
                    "Talk room lifecycle tracking failed.",
                    ex);
            }
        }

        internal TalkRoomReconcileResult ReconcileSuccessfulScan(
            TalkServiceConfiguration configuration,
            IList<TalkRoomTrackingSnapshot> snapshots,
            DateTime nowUtc)
        {
            var result = new TalkRoomReconcileResult();
            if (configuration == null
                || !configuration.IsComplete()
                || snapshots == null)
            {
                return result;
            }

            string accountId;
            try
            {
                accountId =
                    NextcloudUserIdentityService.ResolveCurrentUserId(
                        configuration);
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(
                    LogCategories.Talk,
                    "Talk lifecycle scan retained all rooms because the account could not be verified.",
                    ex);
                return result;
            }
            if (string.IsNullOrWhiteSpace(accountId))
            {
                return result;
            }

            string baseUrl = NormalizeBaseUrl(configuration);
            string login = NormalizeLogin(configuration);
            lock (_syncRoot)
            {
                UpsertPresentSnapshotsLocked(
                    snapshots,
                    baseUrl,
                    login,
                    accountId);
                ConfirmMissingRecordsLocked(
                    snapshots,
                    baseUrl,
                    login,
                    accountId,
                    nowUtc,
                    result);
                SaveLocked();
            }
            if (result.ConfirmedDeletionCount > 0)
            {
                Signal();
            }
            return result;
        }

        internal void QueueUnconditionalDeletion(
            string roomToken,
            bool isEventConversation,
            TalkServiceConfiguration configuration)
        {
            if (string.IsNullOrWhiteSpace(roomToken)
                || configuration == null
                || !configuration.IsComplete())
            {
                return;
            }

            try
            {
                var snapshot = new TalkRoomTrackingSnapshot
                {
                    RoomToken = roomToken.Trim(),
                    IsEventConversation = isEventConversation,
                    Configuration = configuration
                };
                string baseUrl = NormalizeBaseUrl(configuration);
                string login = NormalizeLogin(configuration);
                lock (_syncRoot)
                {
                    TalkRoomLifecycleRecord record =
                        FindSnapshotLocked(
                            snapshot,
                            baseUrl,
                            login,
                            string.Empty);
                    if (record == null)
                    {
                        record = NewRecord(baseUrl, login, string.Empty);
                    }
                    ApplySnapshot(record, snapshot);
                    record.PendingDeletion = true;
                    record.PolicyRequired = false;
                    record.AttemptCount = 0;
                    record.NextAttemptUtc = DateTime.UtcNow;
                    SaveLocked();
                }
                Signal();
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(
                    LogCategories.Talk,
                    "Talk room cleanup could not be queued.",
                    ex);
            }
        }

        internal static bool MatchesAccount(
            TalkRoomLifecycleRecord record,
            string baseUrl,
            string login,
            string accountId)
        {
            if (record == null
                || !string.Equals(
                    record.ServerBaseUrl,
                    baseUrl,
                    StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }
            return !string.IsNullOrWhiteSpace(record.AccountId)
                ? string.Equals(
                    record.AccountId,
                    accountId,
                    StringComparison.Ordinal)
                : string.Equals(
                    record.AccountLogin,
                    login,
                    StringComparison.OrdinalIgnoreCase);
        }

        private void UpsertPresentSnapshotsLocked(
            IList<TalkRoomTrackingSnapshot> snapshots,
            string baseUrl,
            string login,
            string accountId)
        {
            for (int i = 0; i < snapshots.Count; i++)
            {
                TalkRoomTrackingSnapshot snapshot = snapshots[i];
                if (!IsUsable(snapshot)
                    || !string.Equals(
                        NormalizeBaseUrl(snapshot.Configuration),
                        baseUrl,
                        StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                TalkRoomLifecycleRecord record =
                    FindSnapshotLocked(
                        snapshot,
                        baseUrl,
                        login,
                        accountId);
                if (record == null)
                {
                    record = NewRecord(baseUrl, login, accountId);
                }
                else if (string.Equals(
                             record.AccountId,
                             accountId,
                             StringComparison.Ordinal)
                         || (string.IsNullOrWhiteSpace(record.AccountId)
                             && string.Equals(
                                 record.AccountLogin,
                                 login,
                                 StringComparison.OrdinalIgnoreCase)))
                {
                    record.AccountId = accountId;
                    record.AccountLogin = login;
                }
                ApplySnapshot(record, snapshot);
                ResetMissing(record);
            }
        }

        private void ConfirmMissingRecordsLocked(
            IList<TalkRoomTrackingSnapshot> snapshots,
            string baseUrl,
            string login,
            string accountId,
            DateTime nowUtc,
            TalkRoomReconcileResult result)
        {
            for (int i = 0; i < _state.Records.Count; i++)
            {
                TalkRoomLifecycleRecord record = _state.Records[i];
                if (record == null
                    || record.PendingDeletion
                    || !MatchesAccount(
                        record,
                        baseUrl,
                        login,
                        accountId)
                    || IsPresent(record, snapshots, baseUrl))
                {
                    continue;
                }

                if (record.MissingConfirmationCount <= 0)
                {
                    record.MissingConfirmationCount = 1;
                    record.FirstMissingUtc = nowUtc;
                }
                DateTime confirmationUtc =
                    record.FirstMissingUtc.Add(
                        MissingConfirmationDelay);
                if (nowUtc < confirmationUtc)
                {
                    result.NeedsFollowUp = true;
                    if (result.NextConfirmationUtc <= DateTime.MinValue
                        || confirmationUtc < result.NextConfirmationUtc)
                    {
                        result.NextConfirmationUtc = confirmationUtc;
                    }
                    continue;
                }

                record.MissingConfirmationCount++;
                record.PendingDeletion = true;
                record.PolicyRequired = true;
                record.AttemptCount = 0;
                record.NextAttemptUtc = nowUtc;
                result.ConfirmedDeletionCount++;
            }
        }

        private static bool IsPresent(
            TalkRoomLifecycleRecord record,
            IList<TalkRoomTrackingSnapshot> snapshots,
            string baseUrl)
        {
            for (int i = 0; i < snapshots.Count; i++)
            {
                TalkRoomTrackingSnapshot snapshot = snapshots[i];
                if (snapshot == null
                    || snapshot.Configuration == null
                    || !string.Equals(
                        NormalizeBaseUrl(snapshot.Configuration),
                        baseUrl,
                        StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }
                if (!string.IsNullOrWhiteSpace(record.GlobalAppointmentId)
                    && string.Equals(
                        record.GlobalAppointmentId,
                        snapshot.GlobalAppointmentId,
                        StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
                if (!string.IsNullOrWhiteSpace(snapshot.RoomToken)
                    && string.Equals(
                        record.RoomToken,
                        snapshot.RoomToken,
                        StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }
            return false;
        }

        private TalkRoomLifecycleRecord FindSnapshotLocked(
            TalkRoomTrackingSnapshot snapshot,
            string baseUrl,
            string login,
            string accountId)
        {
            TalkRoomLifecycleRecord loginFallback = null;
            for (int i = 0; i < _state.Records.Count; i++)
            {
                TalkRoomLifecycleRecord record = _state.Records[i];
                if (record == null
                    || !string.Equals(
                        record.ServerBaseUrl,
                        baseUrl,
                        StringComparison.OrdinalIgnoreCase)
                    || !HasSameAppointment(record, snapshot))
                {
                    continue;
                }
                if (!string.IsNullOrWhiteSpace(accountId)
                    && string.Equals(
                        record.AccountId,
                        accountId,
                        StringComparison.Ordinal))
                {
                    return record;
                }
                if (string.Equals(
                        record.AccountLogin,
                        login,
                        StringComparison.OrdinalIgnoreCase)
                    && (string.IsNullOrWhiteSpace(accountId)
                        || string.IsNullOrWhiteSpace(record.AccountId)))
                {
                    loginFallback = record;
                }
            }
            return loginFallback;
        }

        private static bool HasSameAppointment(
            TalkRoomLifecycleRecord record,
            TalkRoomTrackingSnapshot snapshot)
        {
            if (!string.IsNullOrWhiteSpace(snapshot.GlobalAppointmentId)
                && string.Equals(
                    record.GlobalAppointmentId,
                    snapshot.GlobalAppointmentId,
                    StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
            if (string.Equals(
                    record.RoomToken,
                    snapshot.RoomToken,
                    StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
            return !string.IsNullOrWhiteSpace(snapshot.EntryId)
                   && string.Equals(
                       record.EntryId,
                       snapshot.EntryId,
                       StringComparison.OrdinalIgnoreCase)
                   && string.Equals(
                       record.StoreId,
                       snapshot.StoreId,
                       StringComparison.OrdinalIgnoreCase);
        }

        private TalkRoomLifecycleRecord NewRecord(
            string baseUrl,
            string login,
            string accountId)
        {
            var record = new TalkRoomLifecycleRecord
            {
                ServerBaseUrl = baseUrl,
                AccountLogin = login,
                AccountId = accountId ?? string.Empty
            };
            _state.Records.Add(record);
            return record;
        }

        private static void ApplySnapshot(
            TalkRoomLifecycleRecord record,
            TalkRoomTrackingSnapshot snapshot)
        {
            record.EntryId = snapshot.EntryId ?? string.Empty;
            record.StoreId = snapshot.StoreId ?? string.Empty;
            record.FolderId = snapshot.FolderId ?? string.Empty;
            record.GlobalAppointmentId =
                snapshot.GlobalAppointmentId ?? string.Empty;
            record.RoomToken = snapshot.RoomToken.Trim();
            record.IsEventConversation = snapshot.IsEventConversation;
        }

        private static void ResetMissing(
            TalkRoomLifecycleRecord record)
        {
            record.MissingConfirmationCount = 0;
            record.FirstMissingUtc = DateTime.MinValue;
            record.PendingDeletion = false;
            record.PolicyRequired = false;
            record.AttemptCount = 0;
            record.NextAttemptUtc = DateTime.MinValue;
        }

        private void OnTimer(object state)
        {
            if (Interlocked.CompareExchange(
                    ref _workerRunning,
                    1,
                    0) != 0)
            {
                return;
            }
            try
            {
                ProcessDueRecords();
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(
                    LogCategories.Talk,
                    "Talk room deletion worker failed.",
                    ex);
            }
            finally
            {
                Interlocked.Exchange(ref _workerRunning, 0);
                ScheduleNext();
            }
        }

        private void ProcessDueRecords()
        {
            List<TalkRoomLifecycleRecord> due =
                GetDueRecords(DateTime.UtcNow);
            if (due.Count == 0)
            {
                return;
            }

            TalkServiceConfiguration configuration =
                BuildConfiguration(_getCurrentSettings());
            string accountId = string.Empty;
            if (configuration != null)
            {
                try
                {
                    accountId =
                        NextcloudUserIdentityService.ResolveCurrentUserId(
                            configuration);
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(
                        LogCategories.Talk,
                        "Talk room deletion retained because the account could not be verified.",
                        ex);
                }
            }

            string baseUrl = NormalizeBaseUrl(configuration);
            string login = NormalizeLogin(configuration);
            for (int i = 0; i < due.Count; i++)
            {
                TalkRoomLifecycleRecord record = due[i];
                if (configuration == null
                    || string.IsNullOrWhiteSpace(accountId)
                    || !MatchesAccount(
                        record,
                        baseUrl,
                        login,
                        accountId))
                {
                    Reschedule(record.Id, TimeSpan.FromHours(1));
                    continue;
                }
                if (record.PolicyRequired
                    && !_isSavedEventDeletionEnabled())
                {
                    Remove(record.Id);
                    continue;
                }
                if (!IsStillPending(record.Id))
                {
                    continue;
                }

                try
                {
                    new TalkService(configuration).DeleteRoom(
                        record.RoomToken,
                        record.IsEventConversation);
                    Remove(record.Id);
                    DiagnosticsLogger.Log(
                        LogCategories.Talk,
                        "Confirmed Talk room deletion completed.");
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(
                        LogCategories.Talk,
                        "Confirmed Talk room deletion failed.",
                        ex);
                    Reschedule(
                        record.Id,
                        GetRetryDelay(record.AttemptCount + 1));
                }
            }
        }

        private List<TalkRoomLifecycleRecord> GetDueRecords(
            DateTime nowUtc)
        {
            var result = new List<TalkRoomLifecycleRecord>();
            lock (_syncRoot)
            {
                for (int i = 0; i < _state.Records.Count; i++)
                {
                    TalkRoomLifecycleRecord record = _state.Records[i];
                    if (record != null
                        && record.PendingDeletion
                        && (record.NextAttemptUtc <= DateTime.MinValue
                            || record.NextAttemptUtc <= nowUtc))
                    {
                        result.Add(CloneForDeletion(record));
                    }
                }
            }
            return result;
        }

        private bool IsStillPending(string recordId)
        {
            lock (_syncRoot)
            {
                TalkRoomLifecycleRecord record = FindByIdLocked(recordId);
                return record != null && record.PendingDeletion;
            }
        }

        private void Reschedule(string recordId, TimeSpan delay)
        {
            lock (_syncRoot)
            {
                TalkRoomLifecycleRecord record = FindByIdLocked(recordId);
                if (record == null || !record.PendingDeletion)
                {
                    return;
                }
                record.AttemptCount++;
                record.NextAttemptUtc = DateTime.UtcNow.Add(delay);
                SaveLocked();
            }
        }

        private void Remove(string recordId)
        {
            lock (_syncRoot)
            {
                TalkRoomLifecycleRecord record = FindByIdLocked(recordId);
                if (record == null)
                {
                    return;
                }
                _state.Records.Remove(record);
                SaveLocked();
            }
        }

        private TalkRoomLifecycleRecord FindByIdLocked(string recordId)
        {
            for (int i = 0; i < _state.Records.Count; i++)
            {
                TalkRoomLifecycleRecord record = _state.Records[i];
                if (record != null
                    && string.Equals(
                        record.Id,
                        recordId,
                        StringComparison.Ordinal))
                {
                    return record;
                }
            }
            return null;
        }

        private void Signal()
        {
            Schedule(TimeSpan.Zero);
        }

        private void ScheduleNext()
        {
            DateTime? earliest = null;
            lock (_syncRoot)
            {
                if (_disposed)
                {
                    return;
                }
                for (int i = 0; i < _state.Records.Count; i++)
                {
                    TalkRoomLifecycleRecord record = _state.Records[i];
                    if (record == null || !record.PendingDeletion)
                    {
                        continue;
                    }
                    DateTime due =
                        record.NextAttemptUtc <= DateTime.MinValue
                            ? DateTime.UtcNow
                            : record.NextAttemptUtc;
                    if (!earliest.HasValue || due < earliest.Value)
                    {
                        earliest = due;
                    }
                }
            }
            if (earliest.HasValue)
            {
                TimeSpan delay = earliest.Value - DateTime.UtcNow;
                Schedule(delay > TimeSpan.Zero ? delay : TimeSpan.Zero);
            }
        }

        private void Schedule(TimeSpan delay)
        {
            lock (_syncRoot)
            {
                if (_disposed)
                {
                    return;
                }
                int milliseconds = delay <= TimeSpan.Zero
                    ? 0
                    : (int)Math.Min(
                        int.MaxValue,
                        Math.Ceiling(delay.TotalMilliseconds));
                _timer.Change(milliseconds, Timeout.Infinite);
            }
        }

        private static TalkServiceConfiguration BuildConfiguration(
            AddinSettings settings)
        {
            if (settings == null)
            {
                return null;
            }
            var configuration = new TalkServiceConfiguration(
                settings.ServerUrl,
                settings.Username,
                settings.AppPassword);
            return configuration.IsComplete() ? configuration : null;
        }

        private static string NormalizeBaseUrl(
            TalkServiceConfiguration configuration)
        {
            return configuration == null
                ? string.Empty
                : configuration.GetNormalizedBaseUrl().TrimEnd('/');
        }

        private static string NormalizeLogin(
            TalkServiceConfiguration configuration)
        {
            return configuration == null
                ? string.Empty
                : (configuration.Username ?? string.Empty).Trim();
        }

        private static bool IsUsable(
            TalkRoomTrackingSnapshot snapshot)
        {
            return snapshot != null
                   && !string.IsNullOrWhiteSpace(snapshot.RoomToken)
                   && snapshot.Configuration != null
                   && snapshot.Configuration.IsComplete();
        }

        private static TimeSpan GetRetryDelay(int attempt)
        {
            if (attempt <= 1)
            {
                return TimeSpan.FromMinutes(1);
            }
            if (attempt == 2)
            {
                return TimeSpan.FromMinutes(5);
            }
            if (attempt == 3)
            {
                return TimeSpan.FromMinutes(15);
            }
            return TimeSpan.FromHours(1);
        }

        private static TalkRoomLifecycleRecord CloneForDeletion(
            TalkRoomLifecycleRecord source)
        {
            return new TalkRoomLifecycleRecord
            {
                Id = source.Id,
                RoomToken = source.RoomToken,
                IsEventConversation = source.IsEventConversation,
                ServerBaseUrl = source.ServerBaseUrl,
                AccountLogin = source.AccountLogin,
                AccountId = source.AccountId,
                PendingDeletion = source.PendingDeletion,
                PolicyRequired = source.PolicyRequired,
                AttemptCount = source.AttemptCount,
                NextAttemptUtc = source.NextAttemptUtc
            };
        }

        private void SaveLocked()
        {
            _store.Save(_state);
        }

        public void Dispose()
        {
            lock (_syncRoot)
            {
                _disposed = true;
                _timer.Change(Timeout.Infinite, Timeout.Infinite);
            }
            _timer.Dispose();
        }
    }
}
