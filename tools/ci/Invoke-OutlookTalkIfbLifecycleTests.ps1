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

$calendarLifecycle = Join-Path $SourceRoot "NextcloudTalkAddIn.TalkCalendarLifecycle.cs"
$calendarMonitor = Join-Path $SourceRoot "Services\TalkCalendarLifecycleMonitor.cs"
$appointmentController = Join-Path $SourceRoot "Controllers\TalkAppointmentController.cs"
$appointmentSyncController = Join-Path $SourceRoot "Controllers\TalkAppointmentController.Sync.cs"
$appointmentSubscription = Join-Path $SourceRoot "NextcloudTalkAddIn.AppointmentSubscription.cs"
$appointmentSync = Join-Path $SourceRoot "NextcloudTalkAddIn.TalkAppointmentSync.cs"
$ifbManager = Join-Path $SourceRoot "Services\FreeBusyManager.cs"
$ifbServer = Join-Path $SourceRoot "Services\FreeBusyServer.cs"
$ifbCache = Join-Path $SourceRoot "Services\IfbAddressBookCache.cs"
$ifbOwnership = Join-Path $SourceRoot "Services\IfbRegistryOwnershipManager.cs"
$talkStore = Join-Path $SourceRoot "Services\TalkRoomLifecycleStore.cs"
$ifbStore = Join-Path $SourceRoot "Services\IfbRegistryStateStore.cs"

Assert-SourceAbsent `
    "Calendar lifecycle never infers deletion from EntryID lookup" `
    $calendarLifecycle `
    'GetItemFromID|IsMapiNotFound'
Assert-SourceContract `
    "Deletion requires a successful all-calendar reconciliation" `
    $calendarLifecycle `
    'TryScanAll[\s\S]*ReconcileSuccessfulScan'
Assert-SourceContract `
    "Calendar lifecycle watches ItemRemove instead of retaining appointment subscriptions" `
    $calendarMonitor `
    'ItemsEvents_ItemRemoveEventHandler'
Assert-SourceContract `
    "Calendar lifecycle enumerates every mounted Outlook store" `
    $calendarMonitor `
    'session\.Stores'
Assert-SourceAbsent `
    "BeforeDelete does not queue a room deletion" `
    $appointmentSubscription `
    'QueueSavedEventRoomDeletion|QueueTrackedSavedEventRoomDeletion'
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

    internal static class LogCategories
    {
        internal const string Talk = "talk";
        internal const string Ifb = "ifb";
    }

    internal static class AppDataPaths
    {
        internal static string EnsureLocalRootDirectory()
        {
            return Path.GetTempPath();
        }
    }
}

namespace NcTalkOutlookAddIn.Settings
{
    internal sealed class AddinSettings
    {
        public string ServerUrl { get; set; }
        public string Username { get; set; }
        public string AppPassword { get; set; }
    }
}

namespace NcTalkOutlookAddIn.Services
{
    internal sealed class TalkServiceConfiguration
    {
        internal TalkServiceConfiguration(
            string serverUrl,
            string username,
            string appPassword)
        {
            ServerUrl = serverUrl ?? string.Empty;
            Username = username ?? string.Empty;
            AppPassword = appPassword ?? string.Empty;
        }

        internal string ServerUrl { get; private set; }
        internal string Username { get; private set; }
        internal string AppPassword { get; private set; }

        internal bool IsComplete()
        {
            return ServerUrl.Length > 0
                && Username.Length > 0
                && AppPassword.Length > 0;
        }

        internal string GetNormalizedBaseUrl()
        {
            return ServerUrl.TrimEnd('/');
        }
    }

    internal static class NextcloudUserIdentityService
    {
        internal static string CanonicalUserId = "alice";

        internal static string ResolveCurrentUserId(
            TalkServiceConfiguration configuration)
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
        TestLifecycleReconciliation();
        TestLifecycleAccountMatching();
        TestSyncCoalescing();
        TestDurableReplacement();

        if (failures > 0)
        {
            return 1;
        }
        Console.WriteLine("All Talk/IFB lifecycle tests passed.");
        return 0;
    }

    private static void TestLifecycleReconciliation()
    {
        string root = Path.Combine(
            Path.GetTempPath(),
            "nc4ol-lifecycle-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        var configuration = NewConfiguration("me@example.org");
        var settings = new NcTalkOutlookAddIn.Settings.AddinSettings
        {
            ServerUrl = configuration.ServerUrl,
            Username = configuration.Username,
            AppPassword = configuration.AppPassword
        };
        var coordinator = new TalkRoomLifecycleCoordinator(
            root,
            "profile",
            () => settings,
            () => true);
        try
        {
            DateTime start = DateTime.UtcNow;
            var original = NewLifecycleSnapshot(
                configuration,
                "global-1",
                "room",
                "entry-1");
            coordinator.ReconcileSuccessfulScan(
                configuration,
                new[] { original },
                start);

            var moved = NewLifecycleSnapshot(
                configuration,
                "global-1",
                string.Empty,
                "entry-2");
            TalkRoomReconcileResult movedResult =
                coordinator.ReconcileSuccessfulScan(
                    configuration,
                    new[] { moved },
                    start.AddSeconds(1));
            Check(
                "GlobalAppointmentID keeps a moved appointment active",
                !movedResult.NeedsFollowUp
                && movedResult.ConfirmedDeletionCount == 0);

            var rekeyed = NewLifecycleSnapshot(
                configuration,
                "global-2",
                "room",
                "entry-3");
            TalkRoomReconcileResult rekeyedResult =
                coordinator.ReconcileSuccessfulScan(
                    configuration,
                    new[] { rekeyed },
                    start.AddSeconds(2));
            Check(
                "Room token keeps a rekeyed appointment active",
                !rekeyedResult.NeedsFollowUp
                && rekeyedResult.ConfirmedDeletionCount == 0);

            TalkRoomReconcileResult firstMissing =
                coordinator.ReconcileSuccessfulScan(
                    configuration,
                    new TalkRoomTrackingSnapshot[0],
                    start.AddSeconds(3));
            Check(
                "First complete absence is only a candidate",
                firstMissing.NeedsFollowUp
                && firstMissing.ConfirmedDeletionCount == 0);

            TalkRoomReconcileResult tooSoon =
                coordinator.ReconcileSuccessfulScan(
                    configuration,
                    new TalkRoomTrackingSnapshot[0],
                    start.AddSeconds(10));
            Check(
                "A repeated scan before the safety interval cannot delete",
                tooSoon.NeedsFollowUp
                && tooSoon.ConfirmedDeletionCount == 0);

            TalkRoomReconcileResult confirmed =
                coordinator.ReconcileSuccessfulScan(
                    configuration,
                    new TalkRoomTrackingSnapshot[0],
                    start.AddSeconds(34));
            Check(
                "A second complete absence after the safety interval confirms deletion",
                !confirmed.NeedsFollowUp
                && confirmed.ConfirmedDeletionCount == 1);
        }
        finally
        {
            coordinator.Dispose();
            Directory.Delete(root, true);
        }
    }

    private static void TestLifecycleAccountMatching()
    {
        var record = new TalkRoomLifecycleRecord
        {
            ServerBaseUrl = "https://cloud.example.org/nextcloud",
            AccountLogin = "old-login",
            AccountId = "alice",
            RoomToken = "room"
        };
        Check(
            "Lifecycle matches canonical UID across login aliases",
            TalkRoomLifecycleCoordinator.MatchesAccount(
                record,
                record.ServerBaseUrl,
                "new-login",
                "alice"));
        Check(
            "Lifecycle rejects another canonical account",
            !TalkRoomLifecycleCoordinator.MatchesAccount(
                record,
                record.ServerBaseUrl,
                "old-login",
                "bob"));
        record.AccountId = string.Empty;
        record.AccountLogin = "me@example.org";
        Check(
            "Unbound lifecycle record requires its exact login",
            TalkRoomLifecycleCoordinator.MatchesAccount(
                record,
                record.ServerBaseUrl,
                "me@example.org",
                "alice"));
    }

    private static TalkServiceConfiguration NewConfiguration(
        string login)
    {
        return new TalkServiceConfiguration(
            "https://cloud.example.org/nextcloud",
            login,
            "secret");
    }

    private static TalkRoomTrackingSnapshot NewLifecycleSnapshot(
        TalkServiceConfiguration configuration,
        string globalId,
        string roomToken,
        string entryId)
    {
        return new TalkRoomTrackingSnapshot
        {
            Configuration = configuration,
            GlobalAppointmentId = globalId,
            RoomToken = roomToken,
            EntryId = entryId,
            StoreId = "store",
            FolderId = "calendar"
        };
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
        (Join-Path $SourceRoot "Services\TalkAppointmentSyncCoordinator.cs"),
        (Join-Path $SourceRoot "Services\TalkRoomLifecycleCoordinator.cs"),
        (Join-Path $SourceRoot "Services\TalkRoomLifecycleStore.cs"),
        (Join-Path $SourceRoot "Models\TalkAppointmentSyncSnapshot.cs"),
        (Join-Path $SourceRoot "Models\TalkRoomLifecycleRecord.cs")
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
