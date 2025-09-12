using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXmlPowerTools;

public static class DocxCleaner
{
    // Removes ALL occurrences of "placeholder" in body, headers, footers,
    // comments, footnotes, endnotes, and text boxes.
    public static void RemoveTextFromDocument(string filePath)
    {
        using var doc = WordprocessingDocument.Open(filePath, true);

        // Build a case-insensitive pattern for the exact text "placeholder".
        var pattern = new Regex(Regex.Escape("placeholder"), RegexOptions.IgnoreCase);

        // Body
        OpenXmlRegex.Replace(doc.MainDocumentPart.Document.Body, pattern, string.Empty, null);

        // Headers / Footers
        foreach (var hp in doc.MainDocumentPart.HeaderParts)
            OpenXmlRegex.Replace(hp.Header, pattern, string.Empty, null);

        foreach (var fp in doc.MainDocumentPart.FooterParts)
            OpenXmlRegex.Replace(fp.Footer, pattern, string.Empty, null);

        // Footnotes / Endnotes
        var fn = doc.MainDocumentPart.FootnotesPart?.Footnotes;
        if (fn != null) OpenXmlRegex.Replace(fn, pattern, string.Empty, null);

        var en = doc.MainDocumentPart.EndnotesPart?.Endnotes;
        if (en != null) OpenXmlRegex.Replace(en, pattern, string.Empty, null);

        // Comments
        var comments = doc.MainDocumentPart.WordprocessingCommentsPart?.Comments;
        if (comments != null) OpenXmlRegex.Replace(comments, pattern, string.Empty, null);

        // Text inside shapes/text boxes (w:txbxContent) in all parts
        void ReplaceInTextBoxes(OpenXmlPartContainer container)
        {
            foreach (var rel in container.Parts)
            {
                var root = rel.OpenXmlPart.RootElement;
                if (root == null) continue;

                foreach (var tx in root.Descendants<TextBoxContent>())
                    OpenXmlRegex.Replace(tx, pattern, string.Empty, null);
            }
        }
        ReplaceInTextBoxes(doc.MainDocumentPart);

        // Save changes
        doc.MainDocumentPart.Document.Save();
        foreach (var hp in doc.MainDocumentPart.HeaderParts) hp.Header.Save();
        foreach (var fp in doc.MainDocumentPart.FooterParts) fp.Footer.Save();
        if (fn != null) doc.MainDocumentPart.FootnotesPart!.Footnotes!.Save();
        if (en != null) doc.MainDocumentPart.EndnotesPart!.Endnotes!.Save();
        if (comments != null) doc.MainDocumentPart.WordprocessingCommentsPart!.Comments!.Save();
    }
}
