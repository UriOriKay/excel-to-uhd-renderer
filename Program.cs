using System;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using Excel = Microsoft.Office.Interop.Excel;
using PdfiumViewer;

internal static class Programm
{
    /// <summary>
    /// Entry poin of the applicaton.
    /// Expects two arguments:
    /// 1) Input Excel file (.xlsx)
    /// 2) Output PDF file (.pdf)
    /// </summary>
    [STAThread]
    private static int Main(string[] args)
    {
        // Validate argument count
        if (args.Length < 2 )
        {
            Console.WriteLine("Usage: <input.xlsx> <output.pdf>");
            return 1;
        }

        // Normaliz paths
        string inputPath = Path.GetFullPath(args[0]);
        string outputPath = Path.GetFullPath(args[1]);

        Console.WriteLine("Input : " + inputPath);
        Console.WriteLine("Output: " + outputPath);

        // Validate input file existence
        if (!File.Exists(inputPath)) 
        {
            Console.WriteLine("Input file not found");
            return 1;
        }

        try
        {
            ConvertExcelToPdf(inputPath, outputPath);
            Console.WriteLine("Conversion successful.");

            // Safety check: ensure PDF was actually created
            if (!File.Exists(outputPath))
                throw new IOException("Excel export reported success, but PDF was not created: " + outputPath);

        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
            return 2;   
        }

        if (args.Length >= 3)
        {
            var pngDir = Path.GetFullPath(args[2]);
            ConvertPdfToPngPages(outputPath, pngDir, dpi: 300);
            Console.WriteLine("PDF -> PNG done: " + pngDir);
            return 0;
        }else
        {
            Console.WriteLine("PNG Path not found.");
            return 2;
        }
    }

    /// <summary>
    /// Converts an Excel workbook to a PDF file using Excel Interop
    /// </summary>
    /// <param name="inputPath"></param>
    /// <param name="outputPath"></param>
    /// <exception cref="InvalidOperationException"></exception>
    private static void ConvertExcelToPdf(string inputPath, string outputPath)
    {
        // Ensure output directory exists
        var outDir = Path.GetDirectoryName(outputPath);
        if (string.IsNullOrWhiteSpace(outDir))
            throw new InvalidOperationException("Output path has no directory: " + outputPath);
        
        Directory.CreateDirectory(outDir);

        // Overwrite existing file
        if (File.Exists(outputPath))
            File.Delete(outputPath);

        Excel.Application? app = null;
        Excel.Workbook? workbook = null;

        try
        {
            // Start Excel in background
            app = new Excel.Application
            {
                Visible = false,
                DisplayAlerts = false
            };

            // Open workbook read-only
            workbook = app.Workbooks.Open(
                inputPath,
                UpdateLinks: 0,
                ReadOnly: true,
                IgnoreReadOnlyRecommended: true
            );

            // Export entire workbook as PDF
            workbook.ExportAsFixedFormat(
                Type: Excel.XlFixedFormatType.xlTypePDF,
                Filename: outputPath,
                Quality: Excel.XlFixedFormatQuality.xlQualityStandard,
                IncludeDocProperties: true,
                IgnorePrintAreas: false, // Respect defined print areas
                OpenAfterPublish: false
            );
        }
        finally
        {
            // CLose workbook safely
            if(workbook != null)
            {
                workbook.Close(false);
                Marshal.FinalReleaseComObject(workbook);
            }

            // Quit Excel safety
            if (app != null)
            {
                app.Quit();
                Marshal.FinalReleaseComObject(app);
            }

            // Force cleanup of Com reference
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }

    private static void ConvertPdfToPngPages(string pdfPath, string outputDir, int dpi = 300)
    {
        if (!File.Exists(pdfPath))
            throw new FileNotFoundException("PDF not found", pdfPath);

        Directory.CreateDirectory(outputDir);

        using var doc = PdfDocument.Load(pdfPath);

        var baseName = Path.GetFileNameWithoutExtension(pdfPath);

        for (int page = 0; page < doc.PageCount; page++)
        {
            // Render: dipX, dpiY, forPrinting
            using var img = doc.Render(page, dpi, dpi, false);

            var fileName = Path.Combine(outputDir, $"{baseName}-{page + 1:00}.png");
            img.Save(fileName, ImageFormat.Png);
        }
    }
}