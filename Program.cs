using System;
using System.IO;
using Spire.Pdf;
using WF = System.Windows.Forms;

internal static class Program
{
    [STAThread]
    private static void Main()
    {
        WF.Application.EnableVisualStyles();

        using var dialog = new WF.OpenFileDialog
        {
            Title = "Select PDF file(s) to convert",
            Filter = "PDF files (*.pdf)|*.pdf",
            Multiselect = true
        };

        if (dialog.ShowDialog() != WF.DialogResult.OK) return;

        foreach (var pdfPath in dialog.FileNames)
        {
            var outPath = GetUniqueOutputPath(pdfPath);
            using var doc = new PdfDocument();
            doc.LoadFromFile(pdfPath);
            doc.SaveToFile(outPath, FileFormat.DOCX);
            Console.WriteLine($"{Path.GetFileName(pdfPath)} → {Path.GetFileName(outPath)}");
        }
    }

    private static string GetUniqueOutputPath(string pdfPath)
    {
        var dir = Path.GetDirectoryName(pdfPath) ?? Environment.CurrentDirectory;
        var name = Path.GetFileNameWithoutExtension(pdfPath);
        var outPath = Path.Combine(dir, name + ".docx");
        for (int i = 1; File.Exists(outPath); i++)
            outPath = Path.Combine(dir, $"{name} ({i}).docx");
        return outPath;
    }
}
