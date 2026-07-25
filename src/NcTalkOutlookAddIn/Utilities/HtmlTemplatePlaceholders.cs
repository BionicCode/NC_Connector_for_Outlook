using AngleSharp.Dom;
using static System.Net.Mime.MediaTypeNames;

namespace NcTalkOutlookAddIn.Utilities
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;

    /// <summary>
    /// Defines a set of placeholder constants used in the backend HTML templates for generating email content.
    /// </summary>
    internal static class HtmlTemplatePlaceholders
    {
        /// <summary>
        /// The placeholder for the URL value in the share link presentation box template.
        /// </summary>
        /// <remarks>
        /// Semantic type: URL value. Rendering: URL/attribute encoding.
        /// </remarks>
        /// <value>The placeholder string for the URL.</value>
        public const string URL = "{URL}";
        /// <summary>
        /// The placeholder for the link label in the share link presentation box template.
        /// </summary>
        /// <remarks>
        /// Semantic type: field label. Rendering: no-break label rendering.
        /// </remarks>
        /// <value>The placeholder string for the link label.</value>
        public const string LINK_LABEL = "{LINK_LABEL}";
        /// <summary>
        /// The placeholder for the call-to-action (CTA) text in the share link presentation box template.
        /// </summary>
        /// <remarks>
        /// Semantic type: CTA text. Rendering: no-break CTA rendering.
        /// </remarks>
        /// <value>The placeholder string for the call-to-action (CTA) text.</value>
        public const string LINK_CTA = "{LINK_CTA}";
        /// <summary>
        /// The placeholder for the expiration date in the share link presentation box template.
        /// </summary>
        /// <remarks>
        /// Semantic type: date/time. Rendering: no-break date rendering.
        /// </remarks>
        /// <value>The placeholder string for the expiration date.</value>
        public const string EXPIRATIONDATE = "{EXPIRATIONDATE}";
        /// <summary>
        /// The placeholder for the password value in the share link presentation box template.
        /// </summary>
        /// <remarks>
        /// Semantic type: value/prose. Rendering: encoded, wrappable.
        /// </remarks>
        /// <value>The placeholder string for the password.</value>
        public const string PASSWORD = "{PASSWORD}";
        /// <summary>
        /// The placeholder for the link introduction in the share link presentation box template.
        /// </summary>
        /// <remarks>
        /// Semantic type: prose. Rendering: encoded, wrappable.
        /// </remarks>
        /// <value>The placeholder string for the link introduction.</value>
        public const string LINK_INTRO = "{LINK_INTRO}";
        /// <summary>
        /// The placeholder for the note in the share link presentation box template.
        /// </summary>
        /// <remarks>
        /// Semantic type: user prose. Rendering: encoded, wrappable.
        /// </remarks>
        /// <value>The placeholder string for the note.</value>
        public const string NOTE = "{NOTE}";
        /// <summary>
        /// The placeholder for the rights in the share link presentation box template.
        /// </summary>
        /// <remarks>
        /// Semantic type: generated HTML fragment. Rendering: trusted generated HTML.
        /// </remarks>
        /// <value>The placeholder string for the rights.</value>
        public const string RIGHTS = "{RIGHTS}";
    }
}
