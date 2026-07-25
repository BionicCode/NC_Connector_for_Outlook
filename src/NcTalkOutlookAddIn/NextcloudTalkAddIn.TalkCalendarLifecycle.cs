// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using NcTalkOutlookAddIn.Controllers;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Services;
using NcTalkOutlookAddIn.Settings;
using NcTalkOutlookAddIn.Utilities;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace NcTalkOutlookAddIn
{
    // Reconciles the compact room index only after a complete all-calendar scan.
    public sealed partial class NextcloudTalkAddIn
    {
        private readonly object _talkCalendarReconciliationSync =
            new object();
        private TalkCalendarLifecycleMonitor
            _talkCalendarLifecycleMonitor;
        private DateTime _talkCalendarReconciliationDueUtc =
            DateTime.MaxValue;
        private int _talkCalendarReconciliationGeneration;

        private bool EnsureTalkCalendarLifecycleMonitor(
            bool runBootstrap)
        {
            if (_talkCalendarLifecycleMonitor != null)
            {
                return _talkCalendarLifecycleMonitor.SetupComplete;
            }
            if (_outlookApplication == null)
            {
                return false;
            }

            try
            {
                _talkCalendarLifecycleMonitor =
                    new TalkCalendarLifecycleMonitor(
                        _outlookApplication,
                        OnTalkCalendarItemChanged,
                        (storeId, folderId) =>
                            ScheduleTalkCalendarReconciliation(
                                TimeSpan.FromSeconds(3)));
                bool ready =
                    _talkCalendarLifecycleMonitor.SetupComplete
                    && _talkCalendarLifecycleMonitor.FolderCount > 0;
                if (ready)
                {
                    ScheduleTalkCalendarReconciliation(
                        TimeSpan.Zero);
                }
                return ready;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(
                    LogCategories.Talk,
                    "Talk calendar lifecycle monitoring could not be initialized.",
                    ex);
                return false;
            }
        }

        private void OnTalkCalendarItemChanged(
            Outlook.AppointmentItem appointment)
        {
            if (appointment == null
                || _talkAppointmentController == null)
            {
                return;
            }
            try
            {
                string token =
                    TalkAppointmentController.GetUserPropertyText(
                        appointment,
                        IcalToken);
                if (string.IsNullOrWhiteSpace(token))
                {
                    return;
                }

                bool lobbyKnown;
                bool lobbyEnabled;
                bool isEventConversation;
                _talkAppointmentController.ResolveRuntimeRoomTraits(
                    appointment,
                    token,
                    false,
                    false,
                    out lobbyKnown,
                    out lobbyEnabled,
                    out isEventConversation);
                TrackTalkRoomAppointment(
                    appointment,
                    token,
                    isEventConversation);
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(
                    LogCategories.Talk,
                    "A changed Talk appointment could not be tracked.",
                    ex);
            }
        }

        internal void NoteTalkRoomDeletionCandidate()
        {
            ScheduleTalkCalendarReconciliation(
                TimeSpan.FromSeconds(3));
        }

        private void ScheduleTalkCalendarReconciliation(
            TimeSpan delay)
        {
            DateTime dueUtc = DateTime.UtcNow.Add(
                delay > TimeSpan.Zero ? delay : TimeSpan.Zero);
            int generation;
            lock (_talkCalendarReconciliationSync)
            {
                if (_talkCalendarLifecycleMonitor == null)
                {
                    return;
                }
                if (_talkCalendarReconciliationDueUtc <= dueUtc)
                {
                    return;
                }
                _talkCalendarReconciliationDueUtc = dueUtc;
                generation =
                    ++_talkCalendarReconciliationGeneration;
            }

            Task reconciliation =
                Task.Delay(
                        delay > TimeSpan.Zero
                            ? delay
                            : TimeSpan.Zero)
                    .ContinueWith(
                        task => RunOnOutlookUiThreadAsync(
                            () => RunTalkCalendarReconciliation(
                                generation)))
                    .Unwrap();
            reconciliation.ContinueWith(
                task => DiagnosticsLogger.LogException(
                    LogCategories.Talk,
                    "Talk calendar reconciliation failed.",
                    task.Exception),
                TaskContinuationOptions.OnlyOnFaulted);
        }

        private void RunTalkCalendarReconciliation(
            int generation)
        {
            lock (_talkCalendarReconciliationSync)
            {
                if (generation
                    != _talkCalendarReconciliationGeneration)
                {
                    return;
                }
                _talkCalendarReconciliationDueUtc =
                    DateTime.MaxValue;
            }

            TalkCalendarLifecycleMonitor monitor =
                _talkCalendarLifecycleMonitor;
            TalkRoomLifecycleCoordinator coordinator =
                _talkRoomLifecycleCoordinator;
            TalkServiceConfiguration configuration =
                BuildCurrentTalkConfiguration();
            if (monitor == null
                || coordinator == null
                || configuration == null)
            {
                return;
            }

            var snapshots =
                new List<TalkRoomTrackingSnapshot>();
            bool complete = monitor.TryScanAll(
                appointment =>
                {
                    string globalId =
                        appointment.GlobalAppointmentID
                        ?? string.Empty;
                    string token =
                        TalkAppointmentController.GetUserPropertyText(
                            appointment,
                            IcalToken)
                        ?? string.Empty;
                    string roomUrl =
                        TalkAppointmentController.GetUserPropertyText(
                            appointment,
                            IcalUrl)
                        ?? string.Empty;
                    if (string.IsNullOrWhiteSpace(token)
                        || !IsRoomUrlForConfiguration(
                            roomUrl,
                            configuration))
                    {
                        snapshots.Add(
                            new TalkRoomTrackingSnapshot
                            {
                                GlobalAppointmentId = globalId,
                                Configuration = configuration
                            });
                        return;
                    }

                    bool isEventConversation =
                        string.Equals(
                            TalkAppointmentController.GetUserPropertyText(
                                appointment,
                                IcalEvent),
                            "event",
                            StringComparison.OrdinalIgnoreCase);
                    TalkRoomTrackingSnapshot snapshot =
                        _talkAppointmentController.CaptureTrackingSnapshot(
                            appointment,
                            token,
                            isEventConversation);
                    if (snapshot == null)
                    {
                        throw new InvalidOperationException(
                            "A Talk appointment snapshot could not be captured.");
                    }
                    snapshots.Add(snapshot);
                });
            if (!complete)
            {
                DiagnosticsLogger.Log(
                    LogCategories.Talk,
                    "Talk lifecycle reconciliation was incomplete; all room records were retained.");
                ScheduleTalkCalendarReconciliation(
                    TimeSpan.FromMinutes(5));
                return;
            }

            Task<TalkRoomReconcileResult> reconcile =
                Task.Run(
                    () => coordinator.ReconcileSuccessfulScan(
                        configuration,
                        snapshots,
                        DateTime.UtcNow));
            Task completion = reconcile.ContinueWith(
                    task => RunOnOutlookUiThreadAsync(
                        () => ApplyTalkCalendarReconcileResult(
                            task.Result)))
                .Unwrap();
            completion.ContinueWith(
                task => DiagnosticsLogger.LogException(
                    LogCategories.Talk,
                    "Talk calendar reconciliation result failed.",
                    task.Exception),
                TaskContinuationOptions.OnlyOnFaulted);
        }

        private void ApplyTalkCalendarReconcileResult(
            TalkRoomReconcileResult result)
        {
            if (result == null || !result.NeedsFollowUp)
            {
                return;
            }
            TimeSpan delay =
                result.NextConfirmationUtc - DateTime.UtcNow;
            ScheduleTalkCalendarReconciliation(
                delay > TimeSpan.Zero
                    ? delay
                    : TimeSpan.Zero);
        }

        private TalkServiceConfiguration
            BuildCurrentTalkConfiguration()
        {
            AddinSettings settings = _currentSettings != null
                ? _currentSettings.Clone()
                : null;
            if (settings == null)
            {
                return null;
            }
            var configuration = new TalkServiceConfiguration(
                settings.ServerUrl,
                settings.Username,
                settings.AppPassword);
            return configuration.IsComplete()
                ? configuration
                : null;
        }

        private static bool IsRoomUrlForConfiguration(
            string roomUrl,
            TalkServiceConfiguration configuration)
        {
            if (string.IsNullOrWhiteSpace(roomUrl))
            {
                return true;
            }
            string baseUrl =
                configuration.GetNormalizedBaseUrl().TrimEnd('/');
            return roomUrl.Trim().StartsWith(
                baseUrl + "/call/",
                StringComparison.OrdinalIgnoreCase);
        }

        private void DisposeTalkCalendarLifecycleMonitor()
        {
            TalkCalendarLifecycleMonitor monitor =
                _talkCalendarLifecycleMonitor;
            _talkCalendarLifecycleMonitor = null;
            lock (_talkCalendarReconciliationSync)
            {
                _talkCalendarReconciliationGeneration++;
                _talkCalendarReconciliationDueUtc =
                    DateTime.MaxValue;
            }
            if (monitor != null)
            {
                monitor.Dispose();
            }
        }
    }
}
