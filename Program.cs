using System;
using System.IO;
using System.Windows.Forms;
using Spire.Pdf;
using Spire.Pdf.FileFormats; // For FileFormat enum

internal static class Program
{
    [STAThread] // Needed for OpenFileDialog
    private static void Main()
    {
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);

        using var dialog = new OpenFileDialog
        {
            Title = "Select PDF file(s) to convert",
            Filter = "PDF files (*.pdf)|*.pdf",
            Multiselect = true,
            CheckFileExists = true,
            CheckPathExists = true
        };

        if (dialog.ShowDialog() != DialogResult.OK)
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
                Console.WriteLine($"✔ Converted: {Path.GetFileName(pdfPath)} -> {Path.GetFileName(outPath)}");
                success++;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✖ Failed: {Path.GetFileName(pdfPath)}\n    {ex.Message}");
                fail++;
            }
        }

        Console.WriteLine($"\nDone. Success: {success}, Failed: {fail}");
    }

    private static string GetUniqueOutputPath(string pdfPath)
    {
        var dir = Path.GetDirectoryName(pdfPath) ?? Environment.CurrentDirectory;
        var nameWithoutExt = Path.GetFileNameWithoutExtension(pdfPath);
        var outPath = Path.Combine(dir, nameWithoutExt + ".docx");

        if (!File.Exists(outPath)) return outPath;

        int i = 1;
        while (true)
        {
            var candidate = Path.Combine(dir, $"{nameWithoutExt} ({i}).docx");
            if (!File.Exists(candidate)) return candidate;
            i++;
        }
    }
}
