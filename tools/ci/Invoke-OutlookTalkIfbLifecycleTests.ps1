Param(
    [string]$ProjectRoot = "."
)

$ErrorActionPreference = "Stop"
$ProjectRoot = (Resolve-Path $ProjectRoot).Path
$SourceRoot = Join-Path $ProjectRoot "src\NcTalkOutlookAddIn"

function Assert-SourceContract {
    param(
        [string]$Name,
        [string]$Path,
        [string]$Pattern
    )

    $text = Get-Content -LiteralPath $Path -Raw
    if ($text -notmatch $Pattern) {
        throw "Source check failed: $Name"
    }
    Write-Host "[OK] $Name"
}

function Assert-SourceAbsent {
    param(
        [string]$Name,
        [string]$Path,
        [string]$Pattern
    )

    $text = Get-Content -LiteralPath $Path -Raw
    if ($text -match $Pattern) {
        throw "Source check failed: $Name"
    }
    Write-Host "[OK] $Name"
}

function Assert-PathAbsent {
    param(
        [string]$Name,
        [string]$Path
    )

    if (Test-Path -LiteralPath $Path) {
        throw "Source check failed: $Name"
    }
    Write-Host "[OK] $Name"
}

$calendarLifecycle = Join-Path $SourceRoot "NextcloudTalkAddIn.TalkCalendarLifecycle.cs"
$calendarMonitor = Join-Path $SourceRoot "Services\TalkCalendarLifecycleMonitor.cs"
$appointmentController = Join-Path $SourceRoot "Controllers\TalkAppointmentController.cs"
$appointmentSyncController = Join-Path $SourceRoot "Controllers\TalkAppointmentController.Sync.cs"
$appointmentSubscription = Join-Path $SourceRoot "NextcloudTalkAddIn.AppointmentSubscription.cs"
$appointmentSync = Join-Path $SourceRoot "NextcloudTalkAddIn.TalkAppointmentSync.cs"
$talkLifecycle = Join-Path $SourceRoot "NextcloudTalkAddIn.TalkRoomLifecycle.cs"
$lifecycle = Join-Path $SourceRoot "NextcloudTalkAddIn.Lifecycle.cs"
$talkCoordinator = Join-Path $SourceRoot "Services\TalkRoomLifecycleCoordinator.cs"
$ifbManager = Join-Path $SourceRoot "Services\FreeBusyManager.cs"
$ifbServer = Join-Path $SourceRoot "Services\FreeBusyServer.cs"
$ifbCache = Join-Path $SourceRoot "Services\IfbAddressBookCache.cs"
$ifbOwnership = Join-Path $SourceRoot "Services\IfbRegistryOwnershipManager.cs"
$talkStore = Join-Path $SourceRoot "Services\TalkRoomLifecycleStore.cs"
$ifbStore = Join-Path $SourceRoot "Services\IfbRegistryStateStore.cs"

Assert-PathAbsent `
    "Global Talk calendar lifecycle partial is removed" `
    $calendarLifecycle
Assert-PathAbsent `
    "Global Talk calendar monitor is removed" `
    $calendarMonitor
foreach ($path in @(
    $appointmentSubscription,
    $talkLifecycle,
    $lifecycle,
    $talkCoordinator,
    $talkStore
)) {
    Assert-SourceAbsent `
        "Targeted Talk lifecycle does not enumerate Outlook stores in $path" `
        $path `
        '(?i)(?:session|application\.Session)\.Stores'
    Assert-SourceAbsent `
        "Targeted Talk lifecycle does not scan calendar item collections in $path" `
        $path `
        'TryScanAll|ReconcileSuccessfulScan'
    Assert-SourceAbsent `
        "Targeted Talk lifecycle does not subscribe to folder item events in $path" `
        $path `
        'ItemsEvents_Item(?:Add|Change|Remove)'
}
Assert-SourceContract `
    "Outlook startup initializes only the pending Talk deletion queue" `
    $lifecycle `
    'InitializeTalkRoomLifecycle\s*\('
Assert-SourceAbsent `
    "Outlook startup does not start calendar reconciliation" `
    $lifecycle `
    'StartTalkRoomLifecycleRecovery|EnsureTalkCalendarLifecycleMonitor'
Assert-SourceContract `
    "Appointment BeforeDelete rejects cancelled deletion and recurrence instances" `
    $appointmentSubscription `
    'OnBeforeDelete[\s\S]*if\s*\(\s*cancel\s*\)[\s\S]*IsRoomDeletionAllowedForRecurrence\s*\(\s*\)[\s\S]*QueueSavedTalkRoomDeletion\s*\(\s*\)'
Assert-SourceContract `
    "Recurring occurrences and exceptions retain the shared Talk room" `
    $appointmentSubscription `
    'OlRecurrenceState\.olApptNotRecurring[\s\S]*OlRecurrenceState\.olApptMaster'
Assert-SourceContract `
    "Saved appointment deletion skips rooms delegated to another user" `
    $appointmentSubscription `
    'QueueSavedTalkRoomDeletion[\s\S]*IsDelegatedToOtherUser[\s\S]*return;[\s\S]*_owner\.QueueSavedTalkRoomDeletion'
Assert-SourceContract `
    "Saved appointment deletion creates a policy-required durable job" `
    $talkLifecycle `
    'QueueSavedTalkRoomDeletion[\s\S]*QueueTalkRoomDeletion\([\s\S]*true\s*\)'
Assert-SourceContract `
    "Unsaved appointment cleanup creates an unconditional durable job" `
    $talkLifecycle `
    'QueueUnsavedTalkRoomDeletion[\s\S]*QueueTalkRoomDeletion\([\s\S]*false\s*\)'
Assert-SourceAbsent `
    "Remote promotion does not leave the room before Outlook persistence" `
    $appointmentSyncController `
    '\.LeaveRoom\s*\('
Assert-SourceContract `
    "Outlook saves the handoff before the background leave call" `
    $appointmentSyncController `
    'appointment\.Save\(\)'
Assert-SourceContract `
    "LeaveRoom runs only from the post-persistence completion path" `
    $appointmentSync `
    'handoffPersisted[\s\S]*\.LeaveRoom\s*\('
Assert-SourceContract `
    "IFB server does not log raw authenticated request URLs" `
    $ifbServer `
    'TryParseAuthorizedEmailPath'
Assert-SourceContract `
    "Outlook IFB URL keeps the attendee domain placeholder" `
    $ifbManager `
    '/freebusy/%NAME%@%SERVER%\.vfb'
Assert-SourceAbsent `
    "Outlook IFB URL does not guess a domain from connector settings" `
    $ifbManager `
    'GuessDefaultDomain'
Assert-SourceAbsent `
    "IFB address lookup has no local-part fallback" `
    $ifbCache `
    'TryResolveEmail|_localPartToEmail'
Assert-SourceContract `
    "IFB server looks up the complete attendee address" `
    $ifbServer `
    'TryGetUid\(\s*_configuration,\s*_cacheHours,\s*email'
Assert-SourceContract `
    "IFB server limits concurrent requests directly" `
    $ifbServer `
    'SemaphoreSlim\s+_requestSlots'
Assert-SourceContract `
    "IFB cache scope includes profile, base URL, and configured username" `
    $ifbCache `
    'BuildScopeFingerprint\(\s*_profileScope,\s*serverBaseUrl,\s*username'
if ((Get-Content -LiteralPath $ifbServer -Raw) -match 'RawUrl') {
    throw "Source check failed: IFB server references RawUrl."
}
Write-Host "[OK] IFB server never references RawUrl"
foreach ($path in @($appointmentController, $appointmentSubscription)) {
    $text = Get-Content -LiteralPath $path -Raw
    if ($text -match 'CreateTalkService\s*\(') {
        throw "Source check failed: appointment callback path creates a synchronous Talk service in $path."
    }
}
Write-Host "[OK] Appointment callback paths contain no synchronous Talk service creation"
Assert-SourceContract `
    "IFB policy keys are checked before user-owned registry writes" `
    $ifbOwnership `
    'ThrowIfPolicyConflicts'
Assert-SourceContract `
    "Talk lifecycle journal uses durable backup replacement" `
    $talkStore `
    'DurableFileReplace\.CommitPreparedFile'
Assert-SourceContract `
    "IFB ownership journal uses durable backup replacement" `
    $ifbStore `
    'DurableFileReplace\.CommitPreparedFile'

$TempRoot = Join-Path ([System.IO.Path]::GetTempPath()) (
    "nc4ol-talk-ifb-tests-" + [Guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Force -Path $TempRoot | Out-Null

try {
    $testSource = Join-Path $TempRoot "TalkIfbLifecycleTests.cs"
    @'
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Services;

namespace NcTalkOutlookAddIn.Utilities
{
    internal static class DiagnosticsLogger
    {
        internal static void Log(string category, string message) { }
        internal static void LogException(
            string category,
            string message,
            Exception ex) { }
    }
}

namespace NcTalkOutlookAddIn.Settings
{
    internal class AddinSettings
    {
        public string ServerUrl { get; set; }
        public string Username { get; set; }
        public string AppPassword { get; set; }
    }
}

namespace NcTalkOutlookAddIn.Services
{
    internal static class NextcloudUserIdentityService
    {
        internal static string CanonicalUserId = "alice";

        internal static string ResolveCurrentUserId(
            TalkServiceConfiguration configuration,
            bool forceRefresh = false)
        {
            return CanonicalUserId;
        }
    }

    internal sealed class TalkService
    {
        internal TalkService(
            TalkServiceConfiguration configuration) { }

        internal void DeleteRoom(
            string roomToken,
            bool isEventConversation) { }
    }
}

internal static class TalkIfbLifecycleTests
{
    private static int failures;

    private static void Check(
        string name,
        bool condition,
        string detail = "")
    {
        if (condition)
        {
            Console.WriteLine("[OK] " + name);
            return;
        }
        failures++;
        Console.Error.WriteLine(
            "[FAIL] "
            + name
            + (detail.Length == 0 ? "" : ": " + detail));
    }

    public static int Main()
    {
        TestLifecycleAccountMatching();
        TestPolicyRequiredDeletion();
        TestUnconditionalDeletion();
        TestDeletionRetryAcrossRestart();
        TestPendingStoreMigration();
        TestSyncCoalescing();
        TestDurableReplacement();

        if (failures > 0)
        {
            return 1;
        }
        Console.WriteLine("All Talk/IFB lifecycle tests passed.");
        return 0;
    }

    private static void TestLifecycleAccountMatching()
    {
        var record = new TalkRoomLifecycleRecord
        {
            ServerBaseUrl =
                "https://cloud.example.org/nextcloud",
            AccountLogin = "old-login",
            AccountId = "alice",
            RoomToken = "room",
            PendingDeletion = true
        };
        Check(
            "Canonical account ID accepts a changed login alias",
            TalkRoomLifecycleCoordinator.MatchesAccount(
                record,
                record.ServerBaseUrl,
                "new-login",
                "alice"));
        Check(
            "Canonical account ID rejects another account",
            !TalkRoomLifecycleCoordinator.MatchesAccount(
                record,
                record.ServerBaseUrl,
                "old-login",
                "bob"));

        record.AccountId = string.Empty;
        record.AccountLogin = "me@example.org";
        Check(
            "Unbound deletion job requires its exact login",
            TalkRoomLifecycleCoordinator.MatchesAccount(
                record,
                record.ServerBaseUrl,
                "me@example.org",
                "alice")
            && !TalkRoomLifecycleCoordinator.MatchesAccount(
                record,
                record.ServerBaseUrl,
                "other@example.org",
                "alice"));
    }

    private static void TestPolicyRequiredDeletion()
    {
        string root = NewTestRoot("policy");
        const string profile = "policy-profile";
        var policyStarted = new ManualResetEventSlim(false);
        var releasePolicy = new ManualResetEventSlim(false);
        int deleteCount = 0;
        var coordinator = new TalkRoomLifecycleCoordinator(
            root,
            profile,
            NewSettings,
            () =>
            {
                policyStarted.Set();
                releasePolicy.Wait(TimeSpan.FromSeconds(5));
                return false;
            },
            configuration => "alice",
            (configuration, token, isEventConversation) =>
                Interlocked.Increment(ref deleteCount));
        try
        {
            Check(
                "Policy-required deletion accepts a complete job",
                coordinator.QueueDeletion(
                    "policy-room",
                    true,
                    NewConfiguration(),
                    true));
            Check(
                "Policy-required deletion reaches the policy gate",
                policyStarted.Wait(TimeSpan.FromSeconds(5)));
            TalkRoomLifecycleState pending =
                LoadLifecycleState(root, profile);
            Check(
                "Policy-required deletion is persisted before evaluation",
                pending.Records.Count == 1
                && pending.Records[0].PolicyRequired
                && pending.Records[0].PendingDeletion
                && pending.Records[0].AccountId == "alice");

            releasePolicy.Set();
            Check(
                "Disabled saved-event deletion removes the pending job",
                WaitUntil(
                    () => ReadCoordinatorRecord(
                        coordinator,
                        "policy-room") == null));
            Check(
                "Policy-off removal is saved to the durable queue",
                LoadLifecycleState(
                    root,
                    profile).Records.Count == 0);
            Check(
                "Disabled saved-event deletion does not call Talk",
                Volatile.Read(ref deleteCount) == 0);
        }
        finally
        {
            releasePolicy.Set();
            coordinator.Dispose();
            Directory.Delete(root, true);
        }
    }

    private static void TestUnconditionalDeletion()
    {
        string root = NewTestRoot("unconditional");
        const string profile = "unconditional-profile";
        var deleted = new ManualResetEventSlim(false);
        var releaseDelete = new ManualResetEventSlim(false);
        int policyCount = 0;
        string deletedToken = string.Empty;
        bool deletedEventRoom = false;
        var coordinator = new TalkRoomLifecycleCoordinator(
            root,
            profile,
            NewSettings,
            () =>
            {
                Interlocked.Increment(ref policyCount);
                return false;
            },
            configuration => "alice",
            (configuration, token, isEventConversation) =>
            {
                deletedToken = token;
                deletedEventRoom = isEventConversation;
                deleted.Set();
                releaseDelete.Wait(TimeSpan.FromSeconds(5));
            });
        try
        {
            Check(
                "Unsaved cleanup accepts a complete job",
                coordinator.QueueDeletion(
                    "unsaved-room",
                    true,
                    NewConfiguration(),
                    false));
            Check(
                "Unsaved cleanup executes without the saved-event policy",
                deleted.Wait(TimeSpan.FromSeconds(5)));
            Check(
                "Unsaved cleanup preserves token and room type",
                deletedToken == "unsaved-room"
                && deletedEventRoom);
            Check(
                "Unsaved cleanup does not evaluate saved-event policy",
                Volatile.Read(ref policyCount) == 0);
            TalkRoomLifecycleRecord pending =
                ReadCoordinatorRecord(
                    coordinator,
                    "unsaved-room");
            Check(
                "Unsaved cleanup is durable before the Talk call completes",
                pending != null
                && pending.PendingDeletion
                && !pending.PolicyRequired);
            releaseDelete.Set();
            Check(
                "Successful unsaved cleanup clears the pending job",
                WaitUntil(
                    () => ReadCoordinatorRecord(
                        coordinator,
                        "unsaved-room") == null));
            Check(
                "Successful unsaved cleanup updates the durable queue",
                LoadLifecycleState(
                    root,
                    profile).Records.Count == 0);
        }
        finally
        {
            releaseDelete.Set();
            coordinator.Dispose();
            Directory.Delete(root, true);
        }
    }

    private static void TestDeletionRetryAcrossRestart()
    {
        string root = NewTestRoot("retry");
        const string profile = "retry-profile";
        var failed = new ManualResetEventSlim(false);
        var releaseFailure = new ManualResetEventSlim(false);
        var firstCoordinator = new TalkRoomLifecycleCoordinator(
            root,
            profile,
            NewSettings,
            () => true,
            configuration => "alice",
            (configuration, token, isEventConversation) =>
            {
                failed.Set();
                releaseFailure.Wait(TimeSpan.FromSeconds(5));
                throw new InvalidOperationException("retry");
            });
        try
        {
            Check(
                "Retry test accepts a complete deletion job",
                firstCoordinator.QueueDeletion(
                    "retry-room",
                    false,
                    NewConfiguration(),
                    true));
            Check(
                "Failed deletion reaches the retry worker",
                failed.Wait(TimeSpan.FromSeconds(5)));
            TalkRoomLifecycleRecord pending =
                ReadCoordinatorRecord(
                    firstCoordinator,
                    "retry-room");
            Check(
                "Deletion job remains pending while Talk is in flight",
                pending != null
                && pending.PendingDeletion
                && pending.AttemptCount == 0);
            releaseFailure.Set();
            Check(
                "Failed deletion remains persisted with retry metadata",
                WaitUntil(
                    () =>
                    {
                        TalkRoomLifecycleRecord record =
                            ReadCoordinatorRecord(
                                firstCoordinator,
                                "retry-room");
                        return record != null
                            && record.AttemptCount == 1
                            && record.NextAttemptUtc
                                > DateTime.UtcNow;
                    }));
            TalkRoomLifecycleState savedRetry =
                LoadLifecycleState(root, profile);
            Check(
                "Retry metadata is written to the durable queue",
                savedRetry.Records.Count == 1
                && savedRetry.Records[0].AttemptCount == 1
                && savedRetry.Records[0].AccountId == "alice");
        }
        finally
        {
            releaseFailure.Set();
            firstCoordinator.Dispose();
        }

        TalkRoomLifecycleState retryState =
            LoadLifecycleState(root, profile);
        retryState.Records[0].NextAttemptUtc = DateTime.MinValue;
        new TalkRoomLifecycleStore(root, profile).Save(retryState);

        var deleted = new ManualResetEventSlim(false);
        var releaseRestartDelete =
            new ManualResetEventSlim(false);
        var restartedCoordinator = new TalkRoomLifecycleCoordinator(
            root,
            profile,
            () => NewSettings("renamed-login"),
            () => true,
            configuration => "alice",
            (configuration, token, isEventConversation) =>
            {
                if (token == "retry-room" && !isEventConversation)
                {
                    deleted.Set();
                    releaseRestartDelete.Wait(
                        TimeSpan.FromSeconds(5));
                }
            });
        try
        {
            restartedCoordinator.StartPendingProcessing();
            Check(
                "Restart matches the canonical account after a login alias change",
                deleted.Wait(TimeSpan.FromSeconds(5)));
            releaseRestartDelete.Set();
            Check(
                "Restarted deletion clears its durable job",
                WaitUntil(
                    () => ReadCoordinatorRecord(
                        restartedCoordinator,
                        "retry-room") == null));
            Check(
                "Restarted deletion updates the durable queue",
                LoadLifecycleState(
                    root,
                    profile).Records.Count == 0);
        }
        finally
        {
            releaseRestartDelete.Set();
            restartedCoordinator.Dispose();
            Directory.Delete(root, true);
        }
    }

    private static void TestPendingStoreMigration()
    {
        string root = NewTestRoot("migration");
        const string profile = "migration-profile";
        var state = new TalkRoomLifecycleState();
        state.Records.Add(
            NewLifecycleRecord("obsolete-room", false));
        state.Records.Add(
            NewLifecycleRecord("pending-room", true));
        new TalkRoomLifecycleStore(root, profile).Save(state);

        var coordinator = new TalkRoomLifecycleCoordinator(
            root,
            profile,
            NewSettings,
            () => true,
            configuration => "alice",
            (configuration, token, isEventConversation) => { });
        try
        {
            TalkRoomLifecycleState migrated =
                LoadLifecycleState(root, profile);
            Check(
                "Startup drops obsolete calendar tracking records",
                migrated.Records.Count == 1
                && migrated.Records[0].RoomToken == "pending-room"
                && migrated.Records[0].PendingDeletion);
        }
        finally
        {
            coordinator.Dispose();
            Directory.Delete(root, true);
        }
    }

    private static TalkRoomLifecycleRecord NewLifecycleRecord(
        string token,
        bool pendingDeletion)
    {
        return new TalkRoomLifecycleRecord
        {
            RoomToken = token,
            ServerBaseUrl =
                "https://cloud.example.org/nextcloud",
            AccountLogin = "me@example.org",
            PendingDeletion = pendingDeletion,
            PolicyRequired = true,
            NextAttemptUtc = DateTime.UtcNow.AddHours(1)
        };
    }

    private static string NewTestRoot(string name)
    {
        string root = Path.Combine(
            Path.GetTempPath(),
            "nc4ol-" + name + "-"
            + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        return root;
    }

    private static NcTalkOutlookAddIn.Settings.AddinSettings
        NewSettings()
    {
        return NewSettings("me@example.org");
    }

    private static NcTalkOutlookAddIn.Settings.AddinSettings
        NewSettings(string username)
    {
        return new NcTalkOutlookAddIn.Settings.AddinSettings
        {
            ServerUrl =
                "https://cloud.example.org/nextcloud",
            Username = username,
            AppPassword = "secret"
        };
    }

    private static TalkServiceConfiguration NewConfiguration()
    {
        NcTalkOutlookAddIn.Settings.AddinSettings settings =
            NewSettings();
        return new TalkServiceConfiguration(
            settings.ServerUrl,
            settings.Username,
            settings.AppPassword);
    }

    private static TalkRoomLifecycleState LoadLifecycleState(
        string root,
        string profile)
    {
        return new TalkRoomLifecycleStore(root, profile).Load();
    }

    private static TalkRoomLifecycleRecord ReadCoordinatorRecord(
        TalkRoomLifecycleCoordinator coordinator,
        string roomToken)
    {
        FieldInfo syncRootField =
            typeof(TalkRoomLifecycleCoordinator).GetField(
                "_syncRoot",
                BindingFlags.NonPublic
                | BindingFlags.Instance);
        FieldInfo stateField =
            typeof(TalkRoomLifecycleCoordinator).GetField(
                "_state",
                BindingFlags.NonPublic
                | BindingFlags.Instance);
        object syncRoot =
            syncRootField.GetValue(coordinator);
        lock (syncRoot)
        {
            TalkRoomLifecycleState state =
                (TalkRoomLifecycleState)stateField.GetValue(
                    coordinator);
            for (int i = 0; i < state.Records.Count; i++)
            {
                TalkRoomLifecycleRecord record =
                    state.Records[i];
                if (record != null
                    && string.Equals(
                        record.RoomToken,
                        roomToken,
                        StringComparison.Ordinal))
                {
                    return new TalkRoomLifecycleRecord
                    {
                        Id = record.Id,
                        RoomToken = record.RoomToken,
                        IsEventConversation =
                            record.IsEventConversation,
                        ServerBaseUrl =
                            record.ServerBaseUrl,
                        AccountLogin = record.AccountLogin,
                        AccountId = record.AccountId,
                        PendingDeletion =
                            record.PendingDeletion,
                        PolicyRequired =
                            record.PolicyRequired,
                        AttemptCount = record.AttemptCount,
                        NextAttemptUtc =
                            record.NextAttemptUtc
                    };
                }
            }
        }
        return null;
    }

    private static bool WaitUntil(Func<bool> condition)
    {
        DateTime deadline =
            DateTime.UtcNow.AddSeconds(5);
        while (DateTime.UtcNow < deadline)
        {
            if (condition())
            {
                return true;
            }
            Thread.Sleep(25);
        }
        return condition();
    }

    private static void TestSyncCoalescing()
    {
        var firstStarted = new ManualResetEventSlim(false);
        var releaseFirst = new ManualResetEventSlim(false);
        var firstCompletionStarted =
            new ManualResetEventSlim(false);
        var releaseFirstCompletion =
            new ManualResetEventSlim(false);
        var completed = new ManualResetEventSlim(false);
        var executed = new List<string>();
        int completionCount = 0;
        var coordinator = new TalkAppointmentSyncCoordinator(
            snapshot =>
            {
                lock (executed)
                {
                    executed.Add(snapshot.RoomName);
                }
                if (snapshot.RoomName == "first")
                {
                    firstStarted.Set();
                    releaseFirst.Wait(TimeSpan.FromSeconds(5));
                }
                return new TalkAppointmentSyncResult(
                    snapshot.RoomToken);
            },
            result =>
            {
                int count =
                    Interlocked.Increment(ref completionCount);
                if (count == 1)
                {
                    firstCompletionStarted.Set();
                    return Task.Run(
                        () => releaseFirstCompletion.Wait(
                            TimeSpan.FromSeconds(5)));
                }
                if (count == 2)
                {
                    completed.Set();
                }
                return Task.FromResult(0);
            });

        coordinator.Queue(NewSnapshot("first"));
        Check(
            "Talk sync worker starts",
            firstStarted.Wait(TimeSpan.FromSeconds(5)));
        coordinator.Queue(NewSnapshot("superseded"));
        coordinator.Queue(NewSnapshot("latest"));
        releaseFirst.Set();
        Check(
            "Talk sync waits for Outlook result persistence",
            firstCompletionStarted.Wait(
                TimeSpan.FromSeconds(5)));
        lock (executed)
        {
            Check(
                "No later remote sync overtakes result persistence",
                executed.Count == 1);
        }
        releaseFirstCompletion.Set();
        Check(
            "Talk sync worker completes coalesced work",
            completed.Wait(TimeSpan.FromSeconds(5)));
        lock (executed)
        {
            Check(
                "Talk sync keeps first in-flight and latest pending snapshot",
                executed.Count == 2
                && executed[0] == "first"
                && executed[1] == "latest",
                string.Join(",", executed.ToArray()));
        }
        coordinator.Dispose();
    }

    private static TalkAppointmentSyncSnapshot NewSnapshot(
        string roomName)
    {
        return new TalkAppointmentSyncSnapshot
        {
            RoomToken = "same-room",
            RoomName = roomName
        };
    }

    private static void TestDurableReplacement()
    {
        string root = Path.Combine(
            Path.GetTempPath(),
            "nc4ol-durable-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try
        {
            string primary = Path.Combine(root, "state.dat");
            string backup = primary + ".bak";
            string prepared = primary + ".tmp";
            File.WriteAllText(primary, "known-good");
            File.WriteAllText(backup, "older");
            File.WriteAllText(prepared, "new-state");
            DurableFileReplace.CommitPreparedFile(
                prepared,
                primary,
                backup);
            Check(
                "Durable replacement writes the new primary",
                File.ReadAllText(primary) == "new-state");
            Check(
                "Durable replacement preserves the prior primary as backup",
                File.ReadAllText(backup) == "known-good");

            MethodInfo fallback =
                typeof(DurableFileReplace).GetMethod(
                    "CommitWithPreservedCopyFallback",
                    BindingFlags.NonPublic
                    | BindingFlags.Static);
            File.WriteAllText(primary, "fallback-old");
            File.WriteAllText(backup, "fallback-older");
            File.WriteAllText(prepared, "fallback-new");
            string preserved = primary + ".preserved";
            File.WriteAllText(preserved, "fallback-old");
            fallback.Invoke(
                null,
                new object[] { prepared, primary, backup, preserved });
            Check(
                "Copy fallback writes the new primary",
                File.ReadAllText(primary) == "fallback-new");
            Check(
                "Copy fallback preserves the prior primary",
                File.ReadAllText(backup) == "fallback-old");
        }
        finally
        {
            Directory.Delete(root, true);
        }
    }
}
'@ | Set-Content -LiteralPath $testSource -Encoding UTF8

    $csc = Join-Path $env:WINDIR "Microsoft.NET\Framework64\v4.0.30319\csc.exe"
    if (-not (Test-Path -LiteralPath $csc)) {
        throw "csc.exe not found at $csc"
    }

    $sources = @(
        $testSource,
        (Join-Path $SourceRoot "Services\DurableFileReplace.cs"),
        (Join-Path $SourceRoot "Services\TalkServiceConfiguration.cs"),
        (Join-Path $SourceRoot "Services\TalkAppointmentSyncCoordinator.cs"),
        (Join-Path $SourceRoot "Services\TalkRoomLifecycleCoordinator.cs"),
        (Join-Path $SourceRoot "Services\TalkRoomLifecycleStore.cs"),
        (Join-Path $SourceRoot "Models\TalkAppointmentSyncSnapshot.cs"),
        (Join-Path $SourceRoot "Models\TalkRoomLifecycleRecord.cs"),
        (Join-Path $SourceRoot "Utilities\AppDataPaths.cs"),
        (Join-Path $SourceRoot "Utilities\LogCategories.cs"),
        (Join-Path $SourceRoot "Utilities\NextcloudUriValidator.cs")
    )
    $exe = Join-Path $TempRoot "TalkIfbLifecycleTests.exe"
    & $csc `
        /nologo `
        /target:exe `
        "/out:$exe" `
        /reference:System.dll `
        /reference:System.Core.dll `
        /reference:System.Security.dll `
        /reference:System.Web.Extensions.dll `
        @sources
    if ($LASTEXITCODE -ne 0) {
        exit $LASTEXITCODE
    }

    & $exe
    if ($LASTEXITCODE -ne 0) {
        exit $LASTEXITCODE
    }
}
finally {
    if (Test-Path -LiteralPath $TempRoot) {
        Remove-Item -LiteralPath $TempRoot -Recurse -Force
    }
}
