using System;
using System.IO;
using Spire.Pdf;
using WF = System.Windows.Forms; // alias to avoid name clashes

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
            Multiselect = true,
            CheckFileExists = true,
            CheckPathExists = true
        };

        if (dialog.ShowDialog() != WF.DialogResult.OK)
        {
            Console.WriteLine("No files selected. Exiting.");
            return;
        }

        int success = 0, fail = 0;

        foreach (var pdfPath in dialog.FileNames)
        {
            try
            {
                var outPath = GetUniqueOutputPath(pdfPath);
                using var doc = new PdfDocument();
                doc.LoadFromFile(pdfPath);
                doc.SaveToFile(outPath, FileFormat.DOCX);
                Console.WriteLine($"✔ {Path.GetFileName(pdfPath)} -> {Path.GetFileName(outPath)}");
                success++;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✖ {Path.GetFileName(pdfPath)}  {ex.Message}");
                fail++;
            }
        }

        Console.WriteLine($"\nDone. Success: {success}, Failed: {fail}");
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
