Param(
    [string]$ProjectRoot = "."
)

$ErrorActionPreference = "Stop"
$ProjectRoot = (Resolve-Path $ProjectRoot).Path

function Assert-Check {
    Param([bool]$Condition, [string]$Message)
    if (-not $Condition) {
        throw $Message
    }
}

function Read-Text {
    Param([string]$RelativePath)
    return Get-Content -Raw -Path (Join-Path $ProjectRoot $RelativePath)
}

$csproj = Read-Text "src\NcTalkOutlookAddIn\NcTalkOutlookAddIn.csproj"
$wxs = Read-Text "installer\Product.wxs"
$vendor = Read-Text "VENDOR.md"

$expectedDlls = @(
    "AngleSharp.dll",
    "AngleSharp.Css.dll",
    "HtmlSanitizer.dll",
    "System.Buffers.dll",
    "System.Collections.Immutable.dll",
    "System.Memory.dll",
    "System.Numerics.Vectors.dll",
    "System.Runtime.CompilerServices.Unsafe.dll",
    "System.Text.Encoding.CodePages.dll"
)

foreach ($dll in $expectedDlls) {
    $relativeVendorPath = "src\NcTalkOutlookAddIn\vendor\htmlsanitizer\$dll"
    Assert-Check (Test-Path (Join-Path $ProjectRoot $relativeVendorPath)) "Vendored dependency missing: $relativeVendorPath"
    Assert-Check ($csproj -match [regex]::Escape("vendor\htmlsanitizer\$dll")) "Project file does not reference vendor DLL: $dll"
    Assert-Check ($wxs -match [regex]::Escape('$' + "(var.BuildOutputDir)" + $dll)) "Installer Product.wxs does not package DLL: $dll"
    Assert-Check ($vendor -match [regex]::Escape($relativeVendorPath.Replace("\", "/")) -or $vendor -match [regex]::Escape($relativeVendorPath)) "VENDOR.md does not document DLL: $dll"
}

function Assert-AssemblyVersion {
    Param(
        [string]$Dll,
        [string]$ExpectedAssemblyVersion,
        [string]$ExpectedFileVersion
    )

    $path = Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\vendor\htmlsanitizer\$Dll"
    $assemblyVersion = [Reflection.AssemblyName]::GetAssemblyName($path).Version.ToString()
    $fileVersion = [Diagnostics.FileVersionInfo]::GetVersionInfo($path).FileVersion
    Assert-Check ($assemblyVersion -eq $ExpectedAssemblyVersion) "$Dll assembly version is $assemblyVersion; expected $ExpectedAssemblyVersion."
    Assert-Check ($fileVersion -eq $ExpectedFileVersion) "$Dll file version is $fileVersion; expected $ExpectedFileVersion."
}

Assert-AssemblyVersion "HtmlSanitizer.dll" "9.0.0.0" "9.0.892.0"
Assert-AssemblyVersion "System.Buffers.dll" "4.0.5.0" "4.600.125.16908"
Assert-AssemblyVersion "System.Collections.Immutable.dll" "10.0.0.0" "10.0.25.52411"
Assert-AssemblyVersion "System.Memory.dll" "4.0.5.0" "4.600.325.20307"
Assert-AssemblyVersion "System.Numerics.Vectors.dll" "4.1.6.0" "4.600.125.16908"
Assert-AssemblyVersion "System.Runtime.CompilerServices.Unsafe.dll" "6.0.3.0" "6.100.225.20307"
Assert-Check ($vendor -match 'HtmlSanitizer/9\.0\.892') "VENDOR.md must document HtmlSanitizer 9.0.892."

Assert-Check ($csproj -match '<Content Include="LICENSE.txt"') "Project file must copy LICENSE.txt to output."
Assert-Check ($csproj -match '<Link>VENDOR\.md</Link>') "Project file must copy VENDOR.md to output."
Assert-Check ($wxs -match 'filLicenseTxt') "Installer must include LICENSE.txt."
Assert-Check ($wxs -match 'filVendorTxt') "Installer must include VENDOR.md."
Assert-Check ($wxs -match 'Bitness="always32"') "Installer must contain explicit 32-bit Outlook registry component."
Assert-Check ($wxs -match 'cmpOutlookAddinReg32') "Installer must contain 32-bit Outlook add-in registration."
Assert-Check ($wxs -match 'cmpOutlookAddinReg"') "Installer must contain 64-bit/default Outlook add-in registration."
Assert-Check ($wxs -match 'http:UrlReservation') "Installer must contain the IFB URL reservation."
Assert-Check (Test-Path (Join-Path $ProjectRoot "installer\assets\nc-connector.ico")) "Installer icon is missing."

Write-Host "Vendor/package check OK: $($expectedDlls.Count) DLLs documented, referenced and packaged."
