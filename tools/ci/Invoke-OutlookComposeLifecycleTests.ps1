Param(
    [string]$ProjectRoot = "."
)

$ErrorActionPreference = "Stop"
$ProjectRoot = (Resolve-Path $ProjectRoot).Path

function Read-Source([string]$RelativePath) {
    return [IO.File]::ReadAllText(
        (Join-Path $ProjectRoot $RelativePath))
}

function Assert-True(
    [string]$Name,
    [bool]$Condition
) {
    if (-not $Condition) {
        throw "[FAIL] $Name"
    }
    Write-Host "[OK] $Name"
}

function Assert-Contains(
    [string]$Name,
    [string]$Source,
    [string]$Expected
) {
    Assert-True $Name $Source.Contains($Expected)
}

function Assert-NotContains(
    [string]$Name,
    [string]$Source,
    [string]$Unexpected
) {
    Assert-True $Name (-not $Source.Contains($Unexpected))
}

$pendingPath = "src\NcTalkOutlookAddIn\Controllers\PendingPasswordDraftController.cs"
$composePath = "src\NcTalkOutlookAddIn\Controllers\ComposeShareLifecycleController.cs"
$trackerPath = "src\NcTalkOutlookAddIn\Controllers\ComposeShareCleanupTracker.cs"
$sendPath = "src\NcTalkOutlookAddIn\NextcloudTalkAddIn.MailComposeSubscription.Send.cs"
$shareCleanupPath = "src\NcTalkOutlookAddIn\NextcloudTalkAddIn.MailComposeSubscription.ShareCleanup.cs"
$subscriptionPath = "src\NcTalkOutlookAddIn\NextcloudTalkAddIn.MailComposeSubscription.cs"
$hooksPath = "src\NcTalkOutlookAddIn\NextcloudTalkAddIn.Hooks.cs"
$fileLinkPath = "src\NcTalkOutlookAddIn\Controllers\FileLinkLaunchController.cs"
$projectPath = "src\NcTalkOutlookAddIn\NcTalkOutlookAddIn.csproj"

$pending = Read-Source $pendingPath
$compose = Read-Source $composePath
$tracker = Read-Source $trackerPath
$send = Read-Source $sendPath
$shareCleanup = Read-Source $shareCleanupPath
$subscription = Read-Source $subscriptionPath
$hooks = Read-Source $hooksPath
$fileLink = Read-Source $fileLinkPath
$project = Read-Source $projectPath

Assert-Contains `
    "Primary send is cancelled when pending drafts cannot be persisted" `
    $send `
    "if (!_owner.TryArmPendingPasswordDrafts("
Assert-Contains `
    "Persistence failure sets Outlook cancel" `
    $send `
    "cancel = true;"
Assert-Contains `
    "DeleteAfterSubmit is rejected because no positive copy can exist" `
    $pending `
    "primary.DeleteAfterSubmit"
Assert-Contains `
    "Primary item receives a unique attempt marker" `
    $pending `
    "WriteProperty(primary, AttemptProperty, attempt);"
Assert-Contains `
    "Password draft is saved before it becomes armed" `
    $pending `
    'WriteProperty(draft, StateProperty, "staging");'
Assert-Contains `
    "Pending payload is protected for the current Windows user" `
    $pending `
    "DataProtectionScope.CurrentUser"

Assert-Contains `
    "Actual SaveSentMessageFolder is captured" `
    $pending `
    "primary.SaveSentMessageFolder"
Assert-Contains `
    "Fallback Sent folder is assigned back to the primary item" `
    $pending `
    "primary.SaveSentMessageFolder = folder;"
Assert-Contains `
    "Recovery resolves the exact persisted Sent folder" `
    $pending `
    "session.GetFolderFromID("
Assert-Contains `
    "Exact Sent folder uses ItemAdd confirmation" `
    $pending `
    "items.ItemAdd += OnSentItemAdded;"
Assert-Contains `
    "Confirmation requires Outlook Sent=true" `
    $pending `
    "mail == null || !ReadSent(mail)"
Assert-Contains `
    "Confirmation also verifies the expected folder" `
    $pending `
    "IsExpectedFolder(attempt, mail)"

$sendMethodStart = $compose.IndexOf(
    "internal bool SendPasswordDraft(",
    [StringComparison]::Ordinal)
$sendMethodEnd = $compose.IndexOf(
    "private static bool ReadSubmittedOrAmbiguous(",
    $sendMethodStart,
    [StringComparison]::Ordinal)
Assert-True `
    "Saved-draft send method is present" `
    ($sendMethodStart -ge 0 -and $sendMethodEnd -gt $sendMethodStart)
$sendMethod = $compose.Substring(
    $sendMethodStart,
    $sendMethodEnd - $sendMethodStart)
Assert-Contains `
    "Automatic delivery sends the saved draft" `
    $sendMethod `
    "((Outlook._MailItem)mail).Send();"
Assert-Contains `
    "Failure displays that same draft" `
    $sendMethod `
    "mail.Display(false);"
Assert-NotContains `
    "Failure path does not create a duplicate draft" `
    $sendMethod `
    "CreateItem("
Assert-Contains `
    "Ambiguous Submitted state suppresses a duplicate fallback" `
    ($compose.Replace("`r`n", "`n")) `
    "catch`n            {`n                return true;"

$insertedIndex = $fileLink.IndexOf(
    "bool inserted =",
    [StringComparison]::Ordinal)
$insertFailureIndex = $fileLink.IndexOf(
    "if (!inserted)",
    $insertedIndex,
    [StringComparison]::Ordinal)
$armIndex = $fileLink.IndexOf(
    "composeSubscription.ArmShareCleanup(",
    $insertFailureIndex,
    [StringComparison]::Ordinal)
$passwordRegistrationIndex = $fileLink.IndexOf(
    "if (registerSeparatePassword)",
    $armIndex,
    [StringComparison]::Ordinal)
Assert-True `
    "Successful FileLink insertion arms compose cleanup before follow-up handling" `
    ($insertedIndex -ge 0 `
        -and $insertFailureIndex -gt $insertedIndex `
        -and $armIndex -gt $insertFailureIndex `
        -and $passwordRegistrationIndex -gt $armIndex)

Assert-Contains `
    "Compose cleanup arms the focused tracker" `
    $shareCleanup `
    "_shareCleanupTracker.Arm(record)"
Assert-Contains `
    "Compose cleanup tracker exposes ReleaseAll" `
    $tracker `
    "internal int ReleaseAll()"
Assert-Contains `
    "Compose cleanup tracker exposes Drain" `
    $tracker `
    "internal List<ComposeShareCleanupRecord> Drain()"

$afterWriteIndex = $shareCleanup.IndexOf(
    "private void OnAfterWrite()",
    [StringComparison]::Ordinal)
$releaseIndex = $shareCleanup.IndexOf(
    "_shareCleanupTracker.ReleaseAll();",
    $afterWriteIndex,
    [StringComparison]::Ordinal)
$unloadIndex = $shareCleanup.IndexOf(
    "private void OnUnload()",
    [StringComparison]::Ordinal)
$inspectorCloseIndex = $shareCleanup.IndexOf(
    "private void OnInspectorClosed()",
    [StringComparison]::Ordinal)
$completionIndex = $shareCleanup.IndexOf(
    "private void CompleteComposeShareCleanup(",
    [StringComparison]::Ordinal)
$drainIndex = $shareCleanup.IndexOf(
    "_shareCleanupTracker.Drain();",
    $completionIndex,
    [StringComparison]::Ordinal)
$disposeIndex = $shareCleanup.IndexOf(
    "Dispose(detachItemEvents);",
    $drainIndex,
    [StringComparison]::Ordinal)
$cleanupQueueIndex = $shareCleanup.IndexOf(
    "_owner.QueueCreatedShareCleanup(",
    $disposeIndex,
    [StringComparison]::Ordinal)
Assert-True `
    "AfterWrite releases shares that Outlook persisted" `
    ($afterWriteIndex -ge 0 `
        -and $releaseIndex -gt $afterWriteIndex `
        -and $releaseIndex -lt $unloadIndex)
Assert-True `
    "Inspector close and inline unload share one captured-state finalizer" `
    ($unloadIndex -ge 0 `
        -and $inspectorCloseIndex -gt $unloadIndex `
        -and $completionIndex -gt $inspectorCloseIndex `
        -and $drainIndex -gt $completionIndex `
        -and $disposeIndex -gt $drainIndex `
        -and $cleanupQueueIndex -gt $disposeIndex)
Assert-NotContains `
    "Share cleanup terminal handlers do not inspect the MailItem" `
    $shareCleanup `
    "_mail"

Assert-Contains `
    "Compose subscription hooks AfterWrite" `
    $subscription `
    "_events.AfterWrite += OnAfterWrite;"
Assert-Contains `
    "Compose subscription hooks Unload" `
    $subscription `
    "_events.Unload += OnUnload;"
Assert-Contains `
    "Compose subscription unhooks AfterWrite" `
    $subscription `
    "_events.AfterWrite -= OnAfterWrite;"
Assert-Contains `
    "Compose subscription unhooks Unload" `
    $subscription `
    "_events.Unload -= OnUnload;"
Assert-Contains `
    "Compose subscription hooks the concrete Inspector close event" `
    $subscription `
    "inspectorEvents.Close += OnInspectorClosed;"
Assert-Contains `
    "Compose subscription unhooks the concrete Inspector close event" `
    $subscription `
    "inspectorEvents.Close -= OnInspectorClosed;"
Assert-Contains `
    "A concrete replacement Inspector rebinds the lifecycle sink" `
    $subscription `
    "ComInteropScope.AreSameObject("
Assert-Contains `
    "Fallback binding does not replace an existing Inspector sink" `
    $subscription `
    "(_inspectorEvents != null && inspector == null)"
Assert-Contains `
    "NewInspector passes the concrete Inspector to the compose subscription" `
    ($hooks.Replace("`r`n", "`n")) `
    "null,`n                        inspector);"

$composeCleanupSources = $subscription + "`n" + $send + "`n" + $shareCleanup
foreach ($obsoleteClosePath in @(
    "_events.Close += OnClose;",
    "ScheduleSurfaceCloseVerification",
    "OnCleanupGraceTimerTick",
    "IsMailComposeSurfaceOpen",
    "_cleanupGraceTimer"
)) {
    Assert-NotContains `
        ("Compose cleanup does not use close polling: " + $obsoleteClosePath) `
        $composeCleanupSources `
        $obsoleteClosePath
}
Assert-NotContains `
    "Compose cleanup does not poll Inspector state" `
    $shareCleanup `
    "IsMailComposeSurfaceOpen"

Assert-NotContains `
    "FileLink launch no longer blocks on current-user lookup" `
    $fileLink `
    "currentUserIdTask"
Assert-Contains `
    "Lifecycle origin remains attached to compose share state" `
    $fileLink `
    "ComposeLifecycleOrigin.Create("

$deleteMethodStart = $compose.IndexOf(
    "internal bool TryDeleteComposeShareFolder(",
    [StringComparison]::Ordinal)
$deleteMethodEnd = $compose.IndexOf(
    "internal void CaptureSeparatePasswordSignatureSnapshot(",
    $deleteMethodStart,
    [StringComparison]::Ordinal)
Assert-True `
    "Compose cleanup method is present" `
    ($deleteMethodStart -ge 0 `
        -and $deleteMethodEnd -gt $deleteMethodStart)
$deleteMethod = $compose.Substring(
    $deleteMethodStart,
    $deleteMethodEnd - $deleteMethodStart)
Assert-Contains `
    "Compose cleanup requires its captured origin" `
    $deleteMethod `
    "entry.Origin == null || !entry.Origin.IsComplete()"
Assert-Contains `
    "Compose cleanup uses its captured origin" `
    $deleteMethod `
    "entry.Origin.ToConfiguration()"
Assert-NotContains `
    "Compose cleanup never falls back to current settings" `
    $deleteMethod `
    "_owner.CurrentSettings"

$removed = @(
    "Models\ComposeLifecycleRecord.cs",
    "Controllers\ComposeLifecycleCoordinator.cs",
    "Controllers\ComposeLifecycleCoordinator.FolderLookup.cs",
    "Controllers\ComposeLifecycleCoordinator.Marker.cs",
    "Controllers\ComposeLifecycleCoordinator.Processing.cs",
    "Controllers\ComposeLifecycleCoordinator.Recovery.cs",
    "Controllers\ComposeLifecycleCoordinator.SentItems.cs",
    "Controllers\SeparatePasswordDispatchController.cs",
    "Controllers\SeparatePasswordDispatchController.Preparation.cs",
    "Controllers\SeparatePasswordDispatchController.Recipients.cs",
    "Controllers\SeparatePasswordDispatchController.Sender.cs",
    "Controllers\SeparatePasswordDispatchController.Signature.cs",
    "Services\ComposeLifecycleJournal.cs"
)
foreach ($relative in $removed) {
    Assert-True `
        ("Legacy source removed: " + $relative) `
        (-not (Test-Path (
            Join-Path `
                $ProjectRoot `
                ("src\NcTalkOutlookAddIn\" + $relative))))
    Assert-NotContains `
        ("Legacy project include removed: " + $relative) `
        $project `
        $relative
}
Assert-Contains `
    "Project includes the focused pending-draft controller" `
    $project `
    'Controllers\PendingPasswordDraftController.cs'
Assert-Contains `
    "Project includes the compact cleanup record" `
    $project `
    'Models\ComposeShareCleanupRecord.cs'
Assert-Contains `
    "Project includes the compose cleanup tracker" `
    $project `
    'Controllers\ComposeShareCleanupTracker.cs'
Assert-Contains `
    "Project includes the compose cleanup subscription partial" `
    $project `
    'NextcloudTalkAddIn.MailComposeSubscription.ShareCleanup.cs'

Write-Host "All Outlook compose lifecycle regression checks passed."
