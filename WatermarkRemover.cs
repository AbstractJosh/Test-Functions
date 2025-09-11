using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools;

public static class DocxCleaner
{
    public static void RemoveTextFromDocument(string filePath, string textToRemove = "placeholder")
    {
        // Load DOCX into WmlDocument
        var doc = new WmlDocument(filePath);

        // Simplify markup (merges runs, cleans redundant tags)
        doc = doc.SimplifyMarkup(new SimplifyMarkupSettings
        {
            RemoveBookmarks = true,
            RemoveComments = true,
            RemoveFieldCodes = false,
            RemoveLastRenderedPageBreak = true,
            RemovePermissions = true,
            RemoveProof = true,
            RemoveRsidInfo = true,
            RemoveSmartTags = true,
            RemoveSoftHyphens = true
        });

        // Replace text everywhere (case-insensitive here, set last param = true for case-sensitive)
        doc = TextReplacer.SearchAndReplace(doc, textToRemove, "", false);

        // Save back to file
        doc.SaveAs(filePath);
    }
}
