Param(
    [string]$ProjectRoot = "."
)

$ErrorActionPreference = "Stop"
$ProjectRoot = (Resolve-Path $ProjectRoot).Path
$VendorDirectory = Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\vendor\htmlsanitizer"
$ConfigurationPath = Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\app.config"
$TempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("nc4ol-binding-tests-" + [Guid]::NewGuid().ToString("N"))

function Assert-Check {
    Param([bool]$Condition, [string]$Message)
    if (-not $Condition) {
        throw $Message
    }
}

function Get-PublicKeyToken {
    Param([Reflection.AssemblyName]$AssemblyName)
    $token = $AssemblyName.GetPublicKeyToken()
    if ($null -eq $token -or $token.Length -eq 0) {
        return "null"
    }
    return (($token | ForEach-Object { $_.ToString("x2") }) -join "")
}

function Normalize-Culture {
    Param([string]$Culture)
    if ([string]::IsNullOrWhiteSpace($Culture) -or $Culture -eq "neutral") {
        return "neutral"
    }
    return $Culture.ToLowerInvariant()
}

Assert-Check (Test-Path -LiteralPath $VendorDirectory) "Vendor directory is missing: $VendorDirectory"
Assert-Check (Test-Path -LiteralPath $ConfigurationPath) "Add-in configuration is missing: $ConfigurationPath"

$deliveredAssemblies = @{}
Get-ChildItem -LiteralPath $VendorDirectory -Filter "*.dll" | Sort-Object Name | ForEach-Object {
    $identity = [Reflection.AssemblyName]::GetAssemblyName($_.FullName)
    Assert-Check (-not $deliveredAssemblies.ContainsKey($identity.Name)) "Duplicate vendored assembly name: $($identity.Name)"
    $deliveredAssemblies[$identity.Name] = [PSCustomObject]@{
        Identity = $identity
        Path = $_.FullName
    }
}
Assert-Check ($deliveredAssemblies.Count -gt 0) "No vendored assemblies found."

[xml]$configuration = Get-Content -LiteralPath $ConfigurationPath -Raw
$namespaceManager = New-Object System.Xml.XmlNamespaceManager($configuration.NameTable)
$namespaceManager.AddNamespace("asm", "urn:schemas-microsoft-com:asm.v1")
$redirectNodes = $configuration.SelectNodes(
    "/configuration/runtime/asm:assemblyBinding/asm:dependentAssembly",
    $namespaceManager)
$redirectNames = @{}

foreach ($dependentAssembly in $redirectNodes) {
    $identityNode = $dependentAssembly.SelectSingleNode("asm:assemblyIdentity", $namespaceManager)
    $redirectNode = $dependentAssembly.SelectSingleNode("asm:bindingRedirect", $namespaceManager)
    Assert-Check ($null -ne $identityNode -and $null -ne $redirectNode) "Malformed dependentAssembly entry in app.config."

    $name = $identityNode.GetAttribute("name")
    Assert-Check (-not [string]::IsNullOrWhiteSpace($name)) "Assembly redirect without a name in app.config."
    Assert-Check (-not $redirectNames.ContainsKey($name)) "Duplicate assembly redirect for $name."
    $redirectNames[$name] = $true

    Assert-Check $deliveredAssemblies.ContainsKey($name) "Assembly redirect target is not vendored: $name."
    $delivered = $deliveredAssemblies[$name].Identity
    $newVersion = New-Object Version($redirectNode.GetAttribute("newVersion"))
    $rangeParts = $redirectNode.GetAttribute("oldVersion").Split("-")
    Assert-Check ($rangeParts.Count -eq 2) "Invalid oldVersion range for $name."
    $rangeStart = New-Object Version($rangeParts[0])
    $rangeEnd = New-Object Version($rangeParts[1])

    Assert-Check ($newVersion -eq $delivered.Version) (
        "Redirect for $name targets $newVersion, but the delivered assembly is $($delivered.Version).")
    Assert-Check ($rangeStart -eq [Version]"0.0.0.0" -and $rangeEnd -eq $newVersion) (
        "Redirect range for $name must be 0.0.0.0-$newVersion.")
    Assert-Check (
        $identityNode.GetAttribute("publicKeyToken").ToLowerInvariant() -eq (Get-PublicKeyToken $delivered)) (
        "Redirect publicKeyToken does not match the delivered $name assembly.")
    Assert-Check (
        (Normalize-Culture $identityNode.GetAttribute("culture")) -eq
            (Normalize-Culture $delivered.CultureName)) (
        "Redirect culture does not match the delivered $name assembly.")
}

New-Item -ItemType Directory -Force -Path $TempRoot | Out-Null
try {
    $probeSource = Join-Path $TempRoot "VendorAssemblyBindingProbe.cs"
    @'
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

public sealed class BindingProbe : MarshalByRefObject
{
    public string Load(string requestedIdentity)
    {
        Assembly loaded = Assembly.Load(requestedIdentity);
        return loaded.GetName().FullName + "\n" + loaded.Location;
    }

    public override object InitializeLifetimeService()
    {
        return null;
    }
}

internal static class VendorAssemblyBindingProbe
{
    private sealed class ReferenceCase
    {
        internal string Source;
        internal AssemblyName Requested;
    }

    public static int Main()
    {
        string baseDirectory = AppDomain.CurrentDomain.BaseDirectory;
        var delivered = Directory.GetFiles(baseDirectory, "*.dll")
            .Select(path => new
            {
                Identity = AssemblyName.GetAssemblyName(path),
                Path = Path.GetFullPath(path)
            })
            .ToDictionary(item => item.Identity.Name, StringComparer.OrdinalIgnoreCase);

        var cases = new List<ReferenceCase>();
        foreach (var source in delivered.Values.OrderBy(item => item.Identity.Name))
        {
            Assembly metadata = Assembly.ReflectionOnlyLoadFrom(source.Path);
            foreach (AssemblyName reference in metadata.GetReferencedAssemblies())
            {
                if (delivered.ContainsKey(reference.Name))
                {
                    cases.Add(new ReferenceCase
                    {
                        Source = source.Identity.Name,
                        Requested = reference
                    });
                }
            }
        }

        if (cases.Count == 0)
        {
            Console.Error.WriteLine("No vendored assembly reference edges were found.");
            return 1;
        }

        int failures = 0;
        foreach (ReferenceCase item in cases
            .OrderBy(item => item.Source)
            .ThenBy(item => item.Requested.FullName))
        {
            AppDomain domain = null;
            try
            {
                var setup = new AppDomainSetup
                {
                    ApplicationBase = baseDirectory,
                    ConfigurationFile = Assembly.GetExecutingAssembly().Location + ".config"
                };
                domain = AppDomain.CreateDomain(
                    "vendor-binding-" + Guid.NewGuid().ToString("N"),
                    null,
                    setup);
                var probe = (BindingProbe)domain.CreateInstanceAndUnwrap(
                    Assembly.GetExecutingAssembly().FullName,
                    typeof(BindingProbe).FullName);
                string[] result = probe.Load(item.Requested.FullName).Split(
                    new[] { '\n' },
                    2);
                AssemblyName actual = new AssemblyName(result[0]);
                var expected = delivered[item.Requested.Name];

                if (!IdentityEquals(actual, expected.Identity) ||
                    result.Length != 2 ||
                    !string.Equals(
                        Path.GetFullPath(result[1]),
                        expected.Path,
                        StringComparison.OrdinalIgnoreCase))
                {
                    failures++;
                    Console.Error.WriteLine(
                        "[FAIL] {0} -> {1}; resolved to {2} at {3}",
                        item.Source,
                        item.Requested.FullName,
                        actual.FullName,
                        result.Length == 2 ? result[1] : "<missing>");
                    continue;
                }

                Console.WriteLine(
                    "[OK] {0} -> {1} resolved to {2}",
                    item.Source,
                    item.Requested.FullName,
                    actual.FullName);
            }
            catch (Exception ex)
            {
                failures++;
                Console.Error.WriteLine(
                    "[FAIL] {0} -> {1}: {2}",
                    item.Source,
                    item.Requested.FullName,
                    ex);
            }
            finally
            {
                if (domain != null)
                {
                    AppDomain.Unload(domain);
                }
            }
        }

        if (failures != 0)
        {
            Console.Error.WriteLine(failures + " vendored assembly binding(s) failed.");
            return 1;
        }

        Console.WriteLine(
            "Vendor assembly binding check OK: {0} reference edge(s) resolved.",
            cases.Count);
        return 0;
    }

    private static bool IdentityEquals(AssemblyName first, AssemblyName second)
    {
        return string.Equals(first.Name, second.Name, StringComparison.OrdinalIgnoreCase)
            && first.Version == second.Version
            && string.Equals(
                NormalizeCulture(first.CultureName),
                NormalizeCulture(second.CultureName),
                StringComparison.OrdinalIgnoreCase)
            && TokensEqual(first.GetPublicKeyToken(), second.GetPublicKeyToken());
    }

    private static string NormalizeCulture(string culture)
    {
        return string.IsNullOrWhiteSpace(culture) ? "neutral" : culture;
    }

    private static bool TokensEqual(byte[] first, byte[] second)
    {
        if (first == null || second == null || first.Length != second.Length)
        {
            return false;
        }
        for (int index = 0; index < first.Length; index++)
        {
            if (first[index] != second[index])
            {
                return false;
            }
        }
        return true;
    }
}
'@ | Set-Content -LiteralPath $probeSource -Encoding UTF8

    $csc = Join-Path $env:WINDIR "Microsoft.NET\Framework64\v4.0.30319\csc.exe"
    Assert-Check (Test-Path -LiteralPath $csc) "csc.exe not found at $csc"

    $probeExecutable = Join-Path $TempRoot "VendorAssemblyBindingProbe.exe"
    & $csc /nologo /target:exe "/out:$probeExecutable" /reference:System.Core.dll $probeSource
    if ($LASTEXITCODE -ne 0) {
        exit $LASTEXITCODE
    }

    Get-ChildItem -LiteralPath $VendorDirectory -Filter "*.dll" |
        Copy-Item -Force -Destination $TempRoot
    Copy-Item -LiteralPath $ConfigurationPath -Destination ($probeExecutable + ".config")

    & $probeExecutable
    if ($LASTEXITCODE -ne 0) {
        exit $LASTEXITCODE
    }
}
finally {
    if (Test-Path -LiteralPath $TempRoot) {
        Remove-Item -LiteralPath $TempRoot -Recurse -Force
    }
}
