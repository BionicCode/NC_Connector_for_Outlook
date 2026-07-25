namespace NcTalkOutlookAddIn.Utilities
{
    using System.Web;

    /// <summary>
    /// Encodes short semantic values that must remain one visual token after
    /// Outlook Classic passes the HTML through the Word compose engine.
    /// </summary>
    /// <remarks>
    /// This class provides methods to encode short semantic values that must remain
    /// as a single visual token in the HTML output, even after being processed by
    /// Outlook's Word-based HTML engine. The encoding ensures that spaces and hyphens are preserved as non-breaking characters,
    /// preventing unwanted line breaks in the rendered output. The encoded values are wrapped in a <nobr> tag with a CSS style to enforce no wrapping.
    /// </remarks>
    internal static class HtmlNoBreakEncoder
    {
        internal static string EncodeFieldLabel(string value)
        {
            return EncodeToken(value);
        }

        internal static string EncodeCallToAction(string value)
        {
            return EncodeToken(value);
        }

        internal static string EncodeDateTime(string value)
        {
            return EncodeToken(value);
        }

        private static string EncodeToken(string value)
        {
            string encoded = HttpUtility.HtmlEncode(value ?? string.Empty);

            // These character-level protections survive Outlook's Word HTML
            // rewriting even when white-space CSS is removed.
            encoded = encoded
                .Replace(" ", "&nbsp;")
                .Replace("-", "&#8209;");

            return "<nobr style=\"white-space: nowrap;\">"
                + encoded
                + "</nobr>";
        }
    }
}
