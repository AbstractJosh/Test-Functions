using System.Text.RegularExpressions;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools;

public static class DocxCleaner
{
    // NuGet: DocumentFormat.OpenXml, OpenXmlPowerTools
    public static void RemoveTextFromDocument(string filePath, string textToRemove = "placeholder", bool caseSensitive = false)
    {
        using var doc = WordprocessingDocument.Open(filePath, true);

        // Optional: normalize content so words aren't split oddly across runs
        MarkupSimplifier.SimplifyMarkup(doc, new SimplifyMarkupSettings
        {
            AcceptRevisions = true,
            RemoveComments = true,
            RemoveRsidInfo = true,
            RemoveSoftHyphens = true,
            RemoveBookmarks = false,
            RemoveFieldCodes = false
        });

        // Build regex (ignore case by default)
        var options = caseSensitive ? RegexOptions.None : RegexOptions.IgnoreCase;
        var pattern = new Regex(Regex.Escape(textToRemove), options);

        // w namespace
        XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        // Local helper to replace in any part
        void ReplaceInPart(OpenXmlPart part)
        {
            if (part == null) return;
            var xDoc = part.GetXDocument();
            if (xDoc?.Root == null) return;

            // Replace inside all text nodes <w:t>
            OpenXmlRegex.Replace(xDoc.Root.Descendants(w + "t"), pattern, string.Empty, null);

            part.PutXDocument();
        }

        // Body
        ReplaceInPart(doc.MainDocumentPart);

        // Headers / Footers
        foreach (var hp in doc.MainDocumentPart.HeaderParts) ReplaceInPart(hp);
        foreach (var fp in doc.MainDocumentPart.FooterParts) ReplaceInPart(fp);

        // Footnotes / Endnotes / Comments
        if (doc.MainDocumentPart.FootnotesPart != null) ReplaceInPart(doc.MainDocumentPart.FootnotesPart);
        if (doc.MainDocumentPart.EndnotesPart  != null) ReplaceInPart(doc.MainDocumentPart.EndnotesPart);
        if (doc.MainDocumentPart.WordprocessingCommentsPart != null) ReplaceInPart(doc.MainDocumentPart.WordprocessingCommentsPart);
    }
}
