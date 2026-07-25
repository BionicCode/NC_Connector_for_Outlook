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
$sendPath = "src\NcTalkOutlookAddIn\NextcloudTalkAddIn.MailComposeSubscription.SendCleanup.cs"
$fileLinkPath = "src\NcTalkOutlookAddIn\Controllers\FileLinkLaunchController.cs"
$projectPath = "src\NcTalkOutlookAddIn\NcTalkOutlookAddIn.csproj"

$pending = Read-Source $pendingPath
$compose = Read-Source $composePath
$send = Read-Source $sendPath
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

Assert-NotContains `
    "Close handling does not infer deletion from folder absence" `
    $send `
    "IsComposeStoredIn"
Assert-NotContains `
    "Close handling does not mark an unknown item discarded" `
    $send `
    "MarkComposeDiscarded"
Assert-Contains `
    "Unknown close state is explicitly retained" `
    $send `
    "unknown state retained without destructive cleanup"

Assert-NotContains `
    "FileLink launch no longer blocks on current-user lookup" `
    $fileLink `
    "currentUserIdTask"
Assert-Contains `
    "Lifecycle origin remains attached to compose share state" `
    $fileLink `
    "ComposeLifecycleOrigin.Create("

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

Write-Host "All Outlook compose lifecycle regression checks passed."
