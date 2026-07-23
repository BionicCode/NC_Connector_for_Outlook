Param(
    [string]$ProjectRoot = "."
)

$ErrorActionPreference = "Stop"
$ProjectRoot = (Resolve-Path $ProjectRoot).Path
$TempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("nc4ol-filelink-tests-" + [Guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Force -Path $TempRoot | Out-Null

try {
    $testSource = Join-Path $TempRoot "OutlookFileLinkRenderingTests.cs"
    @'
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using AngleSharp.Dom;
using AngleSharp.Html.Parser;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.Utilities
{
    internal static class DiagnosticsLogger
    {
        internal static bool IsEnabled { get { return false; } }
        internal static void Log(string category, string message) { }
        internal static void LogException(string category, string message, Exception ex) { }
    }

    internal static class LogCategories
    {
        internal const string Core = "core";
        internal const string FileLink = "filelink";
    }
}

internal static class OutlookFileLinkRenderingTests
{
    private static int failures;

    private static void Check(string name, bool condition, string detail = "")
    {
        if (condition)
        {
            Console.WriteLine("[OK] " + name);
            return;
        }
        failures++;
        Console.Error.WriteLine("[FAIL] " + name + (string.IsNullOrEmpty(detail) ? "" : ": " + detail));
    }

    private static string BuildPermissionsHtml(FileLinkPermissionFlags permissions, string[] labels)
    {
        MethodInfo generator = typeof(FileLinkHtmlBuilder).GetMethod(
            "BuildPermissions",
            BindingFlags.NonPublic | BindingFlags.Static);
        Check("Rights HTML generator is discoverable", generator != null);
        if (generator == null)
        {
            return string.Empty;
        }
        return (string)generator.Invoke(null, new object[]
        {
            permissions,
            labels[0],
            labels[1],
            labels[2],
            labels[3]
        });
    }

    private static Dictionary<string, string> ParseStyle(IElement element)
    {
        var properties = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        string style = element == null ? string.Empty : (element.GetAttribute("style") ?? string.Empty);
        foreach (string declaration in style.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries))
        {
            int separator = declaration.IndexOf(':');
            if (separator < 0)
            {
                continue;
            }
            string name = declaration.Substring(0, separator).Trim();
            string value = Regex.Replace(declaration.Substring(separator + 1).Trim(), @"\s+", " ");
            if (name.Length > 0)
            {
                properties[name] = value;
            }
        }
        return properties;
    }

    private static List<IElement> DirectChildren(IElement parent, string tagName)
    {
        if (parent == null)
        {
            return new List<IElement>();
        }
        return parent.Children
            .Where(child => string.Equals(child.TagName, tagName, StringComparison.OrdinalIgnoreCase))
            .ToList();
    }

    private static IElement SingleDirectChild(string name, IElement parent, string tagName)
    {
        List<IElement> children = DirectChildren(parent, tagName);
        Check(name, children.Count == 1, "count=" + children.Count);
        return children.Count == 1 ? children[0] : null;
    }

    private static void AttributeEquals(string name, IElement element, string attributeName, string expected)
    {
        string actual = element == null ? null : element.GetAttribute(attributeName);
        Check(name, string.Equals(expected, actual, StringComparison.Ordinal), "expected '" + expected + "', got '" + actual + "'");
    }

    private static void StyleEquals(string name, Dictionary<string, string> style, string propertyName, string expected)
    {
        string actual;
        bool present = style.TryGetValue(propertyName, out actual);
        string normalizedExpected = NormalizeCssValue(propertyName, expected);
        string normalizedActual = NormalizeCssValue(propertyName, actual);
        Check(name, present && string.Equals(normalizedExpected, normalizedActual, StringComparison.OrdinalIgnoreCase), "expected '" + expected + "', got '" + (actual ?? "<missing>") + "'");
    }

    private static string NormalizeCssValue(string propertyName, string value)
    {
        string normalized = Regex.Replace(
            value ?? string.Empty,
            @"rgb\(\s*(\d{1,3})\s*,\s*(\d{1,3})\s*,\s*(\d{1,3})\s*\)",
            delegate (Match match)
            {
                return "#"
                    + int.Parse(match.Groups[1].Value).ToString("x2")
                    + int.Parse(match.Groups[2].Value).ToString("x2")
                    + int.Parse(match.Groups[3].Value).ToString("x2");
            },
            RegexOptions.IgnoreCase);
        normalized = Regex.Replace(normalized.Trim(), @"\s+", " ");
        if (!string.Equals(propertyName, "padding", StringComparison.OrdinalIgnoreCase))
        {
            return normalized;
        }
        string[] parts = normalized.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length == 1)
        {
            return string.Join(" ", new[] { parts[0], parts[0], parts[0], parts[0] });
        }
        if (parts.Length == 2)
        {
            return string.Join(" ", new[] { parts[0], parts[1], parts[0], parts[1] });
        }
        if (parts.Length == 3)
        {
            return string.Join(" ", new[] { parts[0], parts[1], parts[2], parts[1] });
        }
        return normalized;
    }

    private static void AssertPresentationTable(string name, IElement table)
    {
        AttributeEquals(name + " role", table, "role", "presentation");
        AttributeEquals(name + " border", table, "border", "0");
        AttributeEquals(name + " cellspacing", table, "cellspacing", "0");
        AttributeEquals(name + " cellpadding", table, "cellpadding", "0");
        Dictionary<string, string> style = ParseStyle(table);
        StyleEquals(name + " border collapse", style, "border-collapse", "collapse");
        StyleEquals(name + " natural width", style, "width", "auto");
        StyleEquals(name + " margin", style, "margin", "0");
    }

    private static int CountOccurrences(string value, string token)
    {
        int count = 0;
        int offset = 0;
        while (!string.IsNullOrEmpty(value) && !string.IsNullOrEmpty(token))
        {
            int index = value.IndexOf(token, offset, StringComparison.Ordinal);
            if (index < 0)
            {
                break;
            }
            count++;
            offset = index + token.Length;
        }
        return count;
    }

    private static void AssertPermissionsHtmlContract(
        string caseName,
        string html,
        string[] labels,
        bool[] enabledStates,
        bool requireEncodedEntities)
    {
        Check(caseName + " avoids flexbox", !Regex.IsMatch(html ?? string.Empty, @"display\s*:\s*(?:inline-)?flex", RegexOptions.IgnoreCase), html);
        Check(caseName + " avoids CSS grid", !Regex.IsMatch(html ?? string.Empty, @"display\s*:\s*(?:inline-)?grid", RegexOptions.IgnoreCase), html);
        if (requireEncodedEntities)
        {
            Check(caseName + " check entity count", CountOccurrences(html, "&#10003;") == enabledStates.Count(enabled => enabled), html);
            Check(caseName + " cross entity count", CountOccurrences(html, "&#10007;") == enabledStates.Count(enabled => !enabled), html);
        }

        var parser = new HtmlParser();
        var document = parser.ParseDocument("<!doctype html><html><body>" + (html ?? string.Empty) + "</body></html>");
        IElement body = document.Body;
        Check(caseName + " outer element count", body != null && body.Children.Count() == 1, html);
        if (body == null || body.Children.Count() != 1)
        {
            return;
        }
        IElement outerTable = body.Children.First();
        Check(caseName + " outer element is a table", string.Equals(outerTable.TagName, "table", StringComparison.OrdinalIgnoreCase), outerTable.OuterHtml);
        Check(caseName + " table count", body.QuerySelectorAll("table").Count() == 5, html);
        Check(caseName + " tbody count", body.QuerySelectorAll("tbody").Count() == 5, html);
        Check(caseName + " row count", body.QuerySelectorAll("tr").Count() == 5, html);
        Check(caseName + " cell count", body.QuerySelectorAll("td").Count() == 12, html);
        AssertPresentationTable(caseName + " outer table", outerTable);

        IElement outerBody = SingleDirectChild(caseName + " outer tbody", outerTable, "tbody");
        IElement outerRow = SingleDirectChild(caseName + " parent row", outerBody, "tr");
        List<IElement> permissionCells = DirectChildren(outerRow, "td");
        Check(caseName + " parent permission cell count", permissionCells.Count == 4, "count=" + permissionCells.Count);
        if (permissionCells.Count != 4)
        {
            return;
        }

        for (int index = 0; index < permissionCells.Count; index++)
        {
            IElement permissionCell = permissionCells[index];
            AttributeEquals(caseName + " item " + (index + 1) + " nowrap", permissionCell, "nowrap", "nowrap");
            AttributeEquals(caseName + " item " + (index + 1) + " valign", permissionCell, "valign", "middle");
            Dictionary<string, string> parentStyle = ParseStyle(permissionCell);
            string expectedPadding = index == permissionCells.Count - 1 ? "0" : "0 12px 0 0";
            StyleEquals(
                caseName + " item " + (index + 1) + " spacing",
                parentStyle,
                "padding",
                expectedPadding);
            if (requireEncodedEntities)
            {
                string actualPadding;
                bool hasPadding = parentStyle.TryGetValue("padding", out actualPadding);
                Check(
                    caseName + " item " + (index + 1) + " direct spacing syntax",
                    hasPadding && string.Equals(expectedPadding, actualPadding, StringComparison.Ordinal),
                    "expected '" + expectedPadding + "', got '" + (actualPadding ?? "<missing>") + "'");
            }
            StyleEquals(caseName + " item " + (index + 1) + " white space", parentStyle, "white-space", "nowrap");
            StyleEquals(caseName + " item " + (index + 1) + " vertical alignment", parentStyle, "vertical-align", "middle");

            IElement nestedTable = SingleDirectChild(caseName + " item " + (index + 1) + " nested table", permissionCell, "table");
            AssertPresentationTable(caseName + " nested table " + (index + 1), nestedTable);
            IElement nestedBody = SingleDirectChild(caseName + " item " + (index + 1) + " nested tbody", nestedTable, "tbody");
            IElement nestedRow = SingleDirectChild(caseName + " item " + (index + 1) + " nested row", nestedBody, "tr");
            List<IElement> iconAndLabelCells = DirectChildren(nestedRow, "td");
            Check(caseName + " item " + (index + 1) + " icon-label cell count", iconAndLabelCells.Count == 2, "count=" + iconAndLabelCells.Count);
            if (iconAndLabelCells.Count != 2)
            {
                continue;
            }

            IElement iconCell = iconAndLabelCells[0];
            string expectedColor = enabledStates[index] ? "#0082c9" : "#c62828";
            string expectedSymbol = enabledStates[index] ? "✓" : "✗";
            AttributeEquals(caseName + " icon " + (index + 1) + " width", iconCell, "width", "14");
            AttributeEquals(caseName + " icon " + (index + 1) + " height", iconCell, "height", "14");
            AttributeEquals(caseName + " icon " + (index + 1) + " align", iconCell, "align", "center");
            AttributeEquals(caseName + " icon " + (index + 1) + " valign", iconCell, "valign", "middle");
            Dictionary<string, string> iconStyle = ParseStyle(iconCell);
            StyleEquals(caseName + " icon " + (index + 1) + " CSS width", iconStyle, "width", "14px");
            StyleEquals(caseName + " icon " + (index + 1) + " CSS height", iconStyle, "height", "14px");
            StyleEquals(caseName + " icon " + (index + 1) + " border", iconStyle, "border", "1px solid " + expectedColor);
            StyleEquals(caseName + " icon " + (index + 1) + " color", iconStyle, "color", expectedColor);
            StyleEquals(caseName + " icon " + (index + 1) + " font size", iconStyle, "font-size", "11px");
            StyleEquals(caseName + " icon " + (index + 1) + " font weight", iconStyle, "font-weight", "700");
            StyleEquals(caseName + " icon " + (index + 1) + " line height", iconStyle, "line-height", "14px");
            Check(caseName + " icon " + (index + 1) + " omits MSO line-height rule", !iconStyle.ContainsKey("mso-line-height-rule"), iconCell.GetAttribute("style"));
            StyleEquals(caseName + " icon " + (index + 1) + " text alignment", iconStyle, "text-align", "center");
            StyleEquals(caseName + " icon " + (index + 1) + " vertical alignment", iconStyle, "vertical-align", "middle");
            Check(caseName + " icon " + (index + 1) + " symbol", string.Equals(iconCell.TextContent.Trim(), expectedSymbol, StringComparison.Ordinal), iconCell.TextContent);

            IElement labelCell = iconAndLabelCells[1];
            AttributeEquals(caseName + " label " + (index + 1) + " nowrap", labelCell, "nowrap", "nowrap");
            AttributeEquals(caseName + " label " + (index + 1) + " valign", labelCell, "valign", "middle");
            Dictionary<string, string> labelStyle = ParseStyle(labelCell);
            StyleEquals(caseName + " label " + (index + 1) + " spacing", labelStyle, "padding-left", "5px");
            Check(caseName + " label " + (index + 1) + " avoids padding shorthand", !labelStyle.ContainsKey("padding"), labelCell.GetAttribute("style"));
            StyleEquals(caseName + " label " + (index + 1) + " white space", labelStyle, "white-space", "nowrap");
            StyleEquals(caseName + " label " + (index + 1) + " font weight", labelStyle, "font-weight", "600");
            StyleEquals(caseName + " label " + (index + 1) + " vertical alignment", labelStyle, "vertical-align", "middle");
            Check(caseName + " label " + (index + 1) + " inherits font family", !labelStyle.ContainsKey("font-family"), labelCell.GetAttribute("style"));
            Check(caseName + " label " + (index + 1) + " inherits font size", !labelStyle.ContainsKey("font-size"), labelCell.GetAttribute("style"));
            Check(caseName + " label " + (index + 1) + " localized text", string.Equals(labelCell.TextContent.Trim(), labels[index], StringComparison.Ordinal), labelCell.TextContent);
        }
    }

    private static void TestPermissionsHtmlContract()
    {
        string[] defaultLabels = { "Read", "Upload", "Modify", "Delete" };
        AssertPermissionsHtmlContract(
            "Read-only Rights HTML",
            BuildPermissionsHtml(FileLinkPermissionFlags.Read, defaultLabels),
            defaultLabels,
            new[] { true, false, false, false },
            true);
        AssertPermissionsHtmlContract(
            "All-enabled Rights HTML",
            BuildPermissionsHtml(FileLinkPermissionFlags.Read | FileLinkPermissionFlags.Create | FileLinkPermissionFlags.Write | FileLinkPermissionFlags.Delete, defaultLabels),
            defaultLabels,
            new[] { true, true, true, true },
            true);
        AssertPermissionsHtmlContract(
            "All-disabled Rights HTML",
            BuildPermissionsHtml(FileLinkPermissionFlags.None, defaultLabels),
            defaultLabels,
            new[] { false, false, false, false },
            true);
        AssertPermissionsHtmlContract(
            "Mixed Rights HTML",
            BuildPermissionsHtml(FileLinkPermissionFlags.Create | FileLinkPermissionFlags.Delete, defaultLabels),
            defaultLabels,
            new[] { false, true, false, true },
            true);

        string[] longLabels =
        {
            "Leseberechtigung für sehr lange Übersetzung",
            "Hochladen und neue Dateien erstellen",
            "Vorhandene Dokumente vollständig bearbeiten",
            "Freigegebene Inhalte dauerhaft löschen"
        };
        AssertPermissionsHtmlContract(
            "Long translated Rights HTML",
            BuildPermissionsHtml(FileLinkPermissionFlags.Read | FileLinkPermissionFlags.Write, longLabels),
            longLabels,
            new[] { true, false, true, false },
            true);

        string[] escapedLabels = { "Read & <inspect>", "Upload", "Modify", "Delete" };
        string escapedHtml = BuildPermissionsHtml(FileLinkPermissionFlags.Read | FileLinkPermissionFlags.Delete, escapedLabels);
        AssertPermissionsHtmlContract(
            "Escaped Rights HTML",
            escapedHtml,
            escapedLabels,
            new[] { true, false, false, true },
            true);
        Check("Rights HTML escapes localized label markup", escapedHtml.Contains("Read &amp; &lt;inspect&gt;") && !escapedHtml.Contains("<inspect>"), escapedHtml);

        BackendPolicyStatus policy = BuildCustomTemplatePolicy("{RIGHTS}");
        FileLinkResult result = BuildResult("https://cloud.example.test/nc/s/AbCd1234", "AbCd1234", string.Empty);
        string sanitizedHtml = FileLinkHtmlBuilder.Build(result, new FileLinkRequest(), "custom", policy);
        AssertPermissionsHtmlContract(
            "Sanitized custom-template Rights HTML",
            sanitizedHtml,
            defaultLabels,
            new[] { true, true, false, false },
            false);
    }

    public static int Main()
    {
        Strings.SetPreferredUiLanguage("en");
        TestPermissionsHtmlContract();
        TestNormalModeUsesNextcloudLinkWording();
        TestAttachmentModeKeepsNextcloudSubpath();
        TestPlainTextKeepsNextcloudSubpath();
        TestAttachmentSharePageTarget();
        TestManualShareIgnoresAttachmentTarget();
        TestInvalidZipUrlFailsVisibly();
        TestCustomTemplateResolvesModeAwareLinkVariables();
        TestBackendEffectiveLanguageLocalizesCustomTemplateCopy();
        TestOlderBackendModeAwareTemplateStillRenders();
        TestLegacyCustomTemplateStillRenders();
        TestSecretLinkLabelHidesLongUrlInHtml();

        if (failures > 0)
        {
            Console.Error.WriteLine(failures + " FileLink rendering test(s) failed.");
            return 1;
        }
        Console.WriteLine("All Outlook FileLink rendering tests passed.");
        return 0;
    }

    private static void TestNormalModeUsesNextcloudLinkWording()
    {
        FileLinkResult result = BuildResult("https://cloud.example.test/nc/s/AbCd1234", "AbCd1234", "Secret!");
        FileLinkRequest request = new FileLinkRequest
        {
            ShareName = "Folder",
            AttachmentMode = false,
            PasswordSeparateEnabled = false,
            NoteEnabled = false,
            Permissions = FileLinkPermissionFlags.Read | FileLinkPermissionFlags.Create
        };
        string html = FileLinkHtmlBuilder.Build(result, request, "en");
        string plainText = FileLinkHtmlBuilder.BuildPlainText(result, request, "en");

        Check("Normal HTML labels the share page as a Nextcloud link", html.Contains(">Nextcloud link<"), html);
        Check("Normal plain text labels the share page as a Nextcloud link", plainText.Contains("Nextcloud link: https://cloud.example.test/nc/s/AbCd1234"), plainText);
        Check("Normal share URL does not gain a ZIP suffix", !html.Contains("/AbCd1234/download") && !plainText.Contains("/AbCd1234/download"));
    }

    private static void TestAttachmentModeKeepsNextcloudSubpath()
    {
        FileLinkResult result = BuildResult("https://cloud.example.test/nc/s/AbCd1234", "AbCd1234", "Secret!");
        FileLinkRequest request = BuildAttachmentRequest();
        string html = FileLinkHtmlBuilder.Build(result, request, "en");

        Check("Attachment ZIP URL keeps /nc subpath in HTML", html.Contains("https://cloud.example.test/nc/s/AbCd1234/download"), html);
        Check("Attachment ZIP URL does not drop /nc subpath in HTML", !html.Contains("https://cloud.example.test/s/AbCd1234/download"), html);
        Check("Attachment HTML labels the link as ZIP download", html.Contains(">ZIP download<"), html);
        Check("Attachment HTML explains ZIP download behavior", html.Contains("Download the shared files as a ZIP archive"), html);
    }

    private static void TestPlainTextKeepsNextcloudSubpath()
    {
        FileLinkResult result = BuildResult("https://cloud.example.test/nc/s/AbCd1234", "AbCd1234", "Secret!");
        FileLinkRequest request = BuildAttachmentRequest();
        string plainText = FileLinkHtmlBuilder.BuildPlainText(result, request, "en");

        Check("Attachment ZIP URL keeps /nc subpath in plain text", plainText.Contains("https://cloud.example.test/nc/s/AbCd1234/download"), plainText);
        Check("Attachment plain text labels the link as ZIP download", plainText.Contains("ZIP download: https://cloud.example.test/nc/s/AbCd1234/download"), plainText);
        Check("Plain text stays plain", !plainText.Contains("<a "), plainText);
    }

    private static void TestAttachmentSharePageTarget()
    {
        FileLinkResult result = BuildResult("https://cloud.example.test/index.php/s/AbCd1234", "AbCd1234", string.Empty);
        FileLinkRequest request = BuildAttachmentRequest(AttachmentLinkTarget.SharePage);
        string html = FileLinkHtmlBuilder.Build(result, request, "en");
        string plainText = FileLinkHtmlBuilder.BuildPlainText(result, request, "en");

        foreach (string output in new[] { html, plainText })
        {
            Check("Attachment share-page target keeps the OCS URL", output.Contains("https://cloud.example.test/index.php/s/AbCd1234"), output);
            Check("Attachment share-page target does not add ZIP suffix", !output.Contains("/AbCd1234/download"), output);
            Check("Attachment share-page target uses share-page wording", output.Contains("Nextcloud link"), output);
            Check("Attachment mode still hides recipient rights", !output.Contains("Your permissions") && !output.Contains("Upload"), output);
        }
    }

    private static void TestManualShareIgnoresAttachmentTarget()
    {
        FileLinkResult result = BuildResult("https://cloud.example.test/nc/s/AbCd1234", "AbCd1234", string.Empty);
        var request = new FileLinkRequest
        {
            AttachmentMode = false,
            AttachmentLinkTarget = AttachmentLinkTarget.ZipDownload,
            Permissions = FileLinkPermissionFlags.Read | FileLinkPermissionFlags.Create
        };
        string plainText = FileLinkHtmlBuilder.BuildPlainText(result, request, "en");

        Check("Manual share always keeps the share-page URL", plainText.Contains("Nextcloud link: https://cloud.example.test/nc/s/AbCd1234"), plainText);
        Check("Manual share ignores ZIP target", !plainText.Contains("/AbCd1234/download"), plainText);
        Check("Manual share keeps rights", plainText.Contains("Your permissions") && plainText.Contains("Upload"), plainText);
    }

    private static void TestInvalidZipUrlFailsVisibly()
    {
        FileLinkRequest request = BuildAttachmentRequest();
        FileLinkResult invalidPath = BuildResult("https://cloud.example.test/index.php/apps/files/", "AbCd1234", string.Empty);
        FileLinkResult tokenMismatch = BuildResult("https://cloud.example.test/s/OtherToken", "AbCd1234", string.Empty);
        FileLinkResult invalidScheme = BuildResult("ftp://cloud.example.test/s/AbCd1234", "AbCd1234", string.Empty);
        FileLinkResult trailingPath = BuildResult("https://cloud.example.test/s/AbCd1234/extra", "AbCd1234", string.Empty);

        Check("ZIP mode rejects a URL without /s/<token>", ThrowsInvalidZipUrl(delegate { FileLinkHtmlBuilder.Build(invalidPath, request, "en"); }));
        Check("ZIP mode rejects a share-token mismatch", ThrowsInvalidZipUrl(delegate { FileLinkHtmlBuilder.BuildPlainText(tokenMismatch, request, "en"); }));
        Check("ZIP mode rejects a non-HTTP(S) URL", ThrowsInvalidZipUrl(delegate { FileLinkHtmlBuilder.Build(invalidScheme, request, "en"); }));
        Check("ZIP mode rejects content after /s/<token>", ThrowsInvalidZipUrl(delegate { FileLinkHtmlBuilder.BuildPlainText(trailingPath, request, "en"); }));
    }

    private static bool ThrowsInvalidZipUrl(Action action)
    {
        try
        {
            action();
            return false;
        }
        catch (InvalidOperationException ex)
        {
            return ex.Message.Contains("ZIP download link could not be created");
        }
    }

    private static void TestCustomTemplateResolvesModeAwareLinkVariables()
    {
        const string template = "<p>{LINK_INTRO}</p><p>{LINK_LABEL}: <a href=\"{URL}\">{URL}</a></p>";
        BackendPolicyStatus policy = BuildCustomTemplatePolicy("<p>Legacy template: {URL}</p>", template);
        FileLinkResult result = BuildResult("https://cloud.example.test/nc/s/AbCd1234", "AbCd1234", string.Empty);

        string normalHtml = FileLinkHtmlBuilder.Build(result, new FileLinkRequest(), "custom", policy);
        string zipHtml = FileLinkHtmlBuilder.Build(result, BuildAttachmentRequest(), "custom", policy);
        string sharePageHtml = FileLinkHtmlBuilder.Build(result, BuildAttachmentRequest(AttachmentLinkTarget.SharePage), "custom", policy);
        string normal = FileLinkHtmlBuilder.BuildPlainText(result, new FileLinkRequest(), "custom", policy);
        string zip = FileLinkHtmlBuilder.BuildPlainText(result, BuildAttachmentRequest(), "custom", policy);

        Check("Custom normal template resolves LINK_INTRO", normal.Contains("Open the Nextcloud link below to view the share."), normal);
        Check("Custom normal template resolves LINK_LABEL", normal.Contains("Nextcloud link: https://cloud.example.test/nc/s/AbCd1234"), normal);
        Check("Custom attachment template resolves ZIP LINK_INTRO", zip.Contains("Download the shared files as a ZIP archive"), zip);
        Check("Custom attachment template resolves ZIP LINK_LABEL", zip.Contains("ZIP download: https://cloud.example.test/nc/s/AbCd1234/download"), zip);
        Check("Custom normal HTML uses the versioned template", normalHtml.Contains("Open the Nextcloud link below to view the share."), normalHtml);
        Check("Custom attachment HTML resolves the versioned template in ZIP mode", zipHtml.Contains("ZIP download"), zipHtml);
        Check("Custom attachment HTML resolves the versioned template in share-page mode", sharePageHtml.Contains("Nextcloud link") && !sharePageHtml.Contains("/download"), sharePageHtml);
        Check("Versioned template takes precedence over compatibility template", !normal.Contains("Legacy template") && !normalHtml.Contains("Legacy template"), normal + normalHtml);
    }

    private static void TestLegacyCustomTemplateStillRenders()
    {
        BackendPolicyStatus policy = BuildCustomTemplatePolicy("<p>Legacy link: {URL}</p>");
        FileLinkResult result = BuildResult("https://cloud.example.test/nc/s/AbCd1234", "AbCd1234", string.Empty);
        string plainText = FileLinkHtmlBuilder.BuildPlainText(result, new FileLinkRequest(), "custom", policy);

        Check("Legacy custom template still resolves its existing URL variable", plainText.Contains("Legacy link: https://cloud.example.test/nc/s/AbCd1234"), plainText);
        Check("Legacy custom template is not forced to contain new variables", !plainText.Contains("LINK_INTRO") && !plainText.Contains("LINK_LABEL"), plainText);
    }

    private static void TestBackendEffectiveLanguageLocalizesCustomTemplateCopy()
    {
        const string template = "<p>{LINK_INTRO}</p><p>{LINK_LABEL}: {URL}</p><p>{PASSWORD}</p><p>{RIGHTS}</p>";
        BackendPolicyStatus policy = BuildCustomTemplatePolicy("<p>Legacy: {URL}</p>", template, "de");
        FileLinkResult result = BuildResult("https://cloud.example.test/nc/s/AbCd1234", "AbCd1234", "Secret!");
        var request = new FileLinkRequest
        {
            PasswordSeparateEnabled = true,
            Permissions = FileLinkPermissionFlags.Read | FileLinkPermissionFlags.Create
        };

        string html = FileLinkHtmlBuilder.Build(result, request, "custom", policy);
        string plainText = FileLinkHtmlBuilder.BuildPlainText(result, request, "custom", policy);

        foreach (string output in new[] { html, plainText })
        {
            Check("Backend template language localizes LINK_INTRO", output.Contains("Öffnen Sie den untenstehenden Nextcloud-Link"), output);
            Check("Backend template language localizes LINK_LABEL", output.Contains("Nextcloud-Link"), output);
            Check("Backend template language localizes separate-password hint", output.Contains("Das Passwort wird in einer separaten E-Mail gesendet."), output);
            Check("Backend template language localizes permission names", output.Contains("Lesen") && output.Contains("Hochladen") && output.Contains("Bearbeiten") && output.Contains("Löschen"), output);
        }
    }

    private static void TestOlderBackendModeAwareTemplateStillRenders()
    {
        const string template = "<p>{LINK_INTRO}</p><p>{LINK_LABEL}: <a href=\"{URL}\">{URL}</a></p>";
        BackendPolicyStatus policy = BuildCustomTemplatePolicy(template);
        FileLinkResult result = BuildResult("https://cloud.example.test/nc/s/AbCd1234", "AbCd1234", string.Empty);
        string plainText = FileLinkHtmlBuilder.BuildPlainText(result, new FileLinkRequest(), "custom", policy);

        Check("Older backend template field still resolves LINK_INTRO", plainText.Contains("Open the Nextcloud link below to view the share."), plainText);
        Check("Older backend template field still resolves LINK_LABEL", plainText.Contains("Nextcloud link: https://cloud.example.test/nc/s/AbCd1234"), plainText);
    }

    private static void TestSecretLinkLabelHidesLongUrlInHtml()
    {
        const string secretUrl = "https://cloud.example.test/index.php/apps/secrets/share/1234567890#VeryLongLocalKey";
        FileLinkResult result = BuildResult("https://cloud.example.test/nc/s/AbCd1234", "AbCd1234", secretUrl);
        string html = FileLinkHtmlBuilder.BuildPasswordOnly(result, "en", null, true);

        Check("Secret password mail uses compact link label", html.Contains(">Secret link<"), html);
        Check("Secret password mail keeps URL in href", html.Contains("href=\"" + secretUrl + "\""), html);
        Check("Secret password mail does not render the long URL as visible text", !html.Contains(">" + secretUrl + "<"), html);
    }

    private static FileLinkResult BuildResult(string shareUrl, string token, string password)
    {
        return new FileLinkResult(
            shareUrl,
            "42",
            token,
            password,
            new DateTime(2026, 7, 7),
            FileLinkPermissionFlags.Read | FileLinkPermissionFlags.Create,
            "Folder",
            "NC Connector/Folder");
    }

    private static FileLinkRequest BuildAttachmentRequest(AttachmentLinkTarget target = AttachmentLinkTarget.ZipDownload)
    {
        return new FileLinkRequest
        {
            ShareName = "Folder",
            AttachmentMode = true,
            AttachmentLinkTarget = target,
            PasswordSeparateEnabled = false,
            NoteEnabled = false,
            Permissions = FileLinkPermissionFlags.Read | FileLinkPermissionFlags.Create
        };
    }

    private static BackendPolicyStatus BuildCustomTemplatePolicy(string template, string versionedTemplate = null, string effectiveLanguage = null)
    {
        var sharePolicy = new Dictionary<string, object>
        {
            { "share_html_block_template", template }
        };
        if (!string.IsNullOrWhiteSpace(versionedTemplate))
        {
            sharePolicy.Add("share_html_block_template_v2", versionedTemplate);
        }
        if (!string.IsNullOrWhiteSpace(effectiveLanguage))
        {
            sharePolicy.Add("share_html_block_effective_language", effectiveLanguage);
        }
        var empty = new Dictionary<string, object>();
        return new BackendPolicyStatus(
            true,
            true,
            true,
            false,
            string.Empty,
            "policy",
            string.Empty,
            true,
            true,
            "active",
            sharePolicy,
            empty,
            empty,
            empty,
            empty,
            empty);
    }
}
'@ | Set-Content -Path $testSource -Encoding UTF8

    $csc = Join-Path $env:WINDIR "Microsoft.NET\Framework64\v4.0.30319\csc.exe"
    if (-not (Test-Path $csc)) {
        throw "csc.exe not found at $csc"
    }

    $vendorDir = Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\vendor\htmlsanitizer"
    $sources = @(
        $testSource,
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Models\BackendPolicyStatus.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Models\AttachmentLinkTargetPolicy.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Models\FileLinkPermissions.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Models\FileLinkSelection.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Models\FileLinkRequest.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Models\FileLinkResult.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Models\SharePasswordDeliveryMode.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Utilities\BrandingAssets.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Utilities\FileLinkHtmlBuilder.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Utilities\HtmlTemplateSanitizer.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Utilities\HtmlToPlainTextConverter.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Utilities\Strings.cs")
    )
    $references = @(
        "/reference:System.dll",
        "/reference:System.Core.dll",
        "/reference:System.Drawing.dll",
        "/reference:System.Web.dll",
        "/reference:System.Web.Extensions.dll"
    )
    Get-ChildItem -Path $vendorDir -Filter "*.dll" | ForEach-Object {
        $references += "/reference:$($_.FullName)"
    }
    $resources = @(
        "/resource:$((Resolve-Path (Join-Path $ProjectRoot 'src\NcTalkOutlookAddIn\Resources\_locales\en\messages.json')).Path),OutlookFileLinkRenderingTests.Resources._locales.en.messages.json",
        "/resource:$((Resolve-Path (Join-Path $ProjectRoot 'src\NcTalkOutlookAddIn\Resources\_locales\de\messages.json')).Path),OutlookFileLinkRenderingTests.Resources._locales.de.messages.json"
    )

    $exe = Join-Path $TempRoot "OutlookFileLinkRenderingTests.exe"
    & $csc /nologo /nowarn:1702 /target:exe "/out:$exe" @references @resources @sources
    if ($LASTEXITCODE -ne 0) {
        exit $LASTEXITCODE
    }

    Get-ChildItem -Path $vendorDir -Filter "*.dll" | Copy-Item -Force -Destination $TempRoot

    & $exe
    if ($LASTEXITCODE -ne 0) {
        exit $LASTEXITCODE
    }
}
finally {
    if (Test-Path $TempRoot) {
        Remove-Item -LiteralPath $TempRoot -Recurse -Force
    }
}
