using System;
using System.Drawing;
using System.Drawing.Drawing2D;
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
            ConvertPdfToPngPages(outputPath, pngDir, dpi: 400);
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



    /// <summary>
    /// Renders each page of a PDF document into a separate PNG image file.
    /// </summary>
    /// <param name="pdfPath">
    /// Absolute or ralative path to the surce PDF file.
    /// The file must exist and be reable
    /// </param>
    /// <param name="outputDir">
    /// Directory where the generated PNG files will be written.
    /// The directory will be created if it does not exist
    /// </param>
    /// <param name="dpi">
    /// Rendering resolution in dots per inch
    /// Higher values increase image quality and file size
    /// Default is 300 DPI.
    /// </param>
    /// <exception cref="FileNotFoundException">
    /// Thrown when the specified PDF file does not exist.
    /// </exception>
    /// <remarks>
    /// - One Png file is generated per PDF page.
    /// - File names are based on the PDF file name with a sequential suffix
    ///   (e.g. "document-01.png").
    /// - Uses PdfiumViewer for rendering (native PDFium dependency).
    /// - Only supported on Windows due to System.Drawing and PDFium usage.
    /// </remarks>
    /// <return>
    /// Nothing. PNG files are written to disk as a side effect.
    /// </return>
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

            using var srcBmp = new Bitmap(img.Width, img.Height, PixelFormat.Format32bppArgb);
            using (var g = Graphics.FromImage(srcBmp))
            {
                g.DrawImage(img, 0, 0, img.Width, img.Height);
            }

            using var overlay = ToWhiteTextTransparentBackground(srcBmp, backgroundCutoff: 245, gamma: 0.85f);


            var outFile = Path.Combine(outputDir, $"{baseName}-{page + 1:00}.png");
            SaveUhdCanvasWithCenteredImage(overlay, outFile, 3840, 2160);
        }
    }

    private static void SaveUhdCanvasWithCenteredImage(Image overlay, string outputFile, int targetW = 3840, int targetH =2160)
    {
        using var uhd = new Bitmap(targetW, targetH, PixelFormat.Format32bppArgb);

        using (var g = Graphics.FromImage(uhd))
        {
            g.Clear(Color.Black);

            //scaling
            float scale = Math.Min((float)targetW / overlay.Width, (float)targetH / overlay.Height);

            int drawW = (int)Math.Round(overlay.Width * scale);
            int drawH = (int)Math.Round(overlay.Height * scale);

            int x = (targetW - drawW) / 2;
            int y = (targetH - drawH) / 2;

            g.CompositingMode = CompositingMode.SourceOver;
            g.InterpolationMode = InterpolationMode.HighQualityBicubic;
            g.PixelOffsetMode = PixelOffsetMode.HighQuality;
            g.SmoothingMode = SmoothingMode.HighQuality;

            g.DrawImage(overlay, new Rectangle(x, y, drawW, drawH));

        }

        uhd.Save(outputFile, ImageFormat.Png);
    }

    private static Bitmap ToWhiteTextTransparentBackground(Bitmap input, byte backgroundCutoff = 245, float gamma = 1.0f)
    {
        var output = new Bitmap(input.Width, input.Height, PixelFormat.Format32bppArgb);

        var rect = new Rectangle(0, 0, input.Width, input.Height);
        var srcData = input.LockBits(rect, ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
        var dstData = output.LockBits(rect, ImageLockMode.WriteOnly, PixelFormat.Format32bppArgb); // Fehler ---

        try
        {
            int bytes = Math.Abs(srcData.Stride) * srcData.Height;
            byte[] src = new byte[bytes];
            byte[] dst = new byte[bytes];

            Marshal.Copy(srcData.Scan0, src, 0, bytes);

            for (int i = 0; i < bytes; i += 4)
            {
                byte b = src[i + 0];
                byte g = src[i + 1];
                byte r = src[i + 2];

                int lum = (int)(0.2126 * r + 0.7152 * g + 0.0722 * b);

                if (lum >= backgroundCutoff)
                {
                    dst[i + 0] = 255; // white
                    dst[i + 1] = 255;
                    dst[i + 2] = 255;
                    dst[i + 3] = 0;   // alpha
                    continue;
                }

                // Alpha out of "Dark": black(0) => alpha 255, white(255) => alpha 0
                float a = 255f - lum;

                // optional: gamma für stärkere/schwächere Kanten
                if (gamma != 1.0f)
                {
                    float norm = a / 255f;
                    norm = (float)Math.Pow(norm, gamma);
                    a = norm * 255f;
                }

                byte alpha = (byte)Math.Clamp((int)Math.Round(a), 0, 255);

                dst[i + 0] = 255;   // white
                dst[i + 1] = 255;
                dst[i + 2] = 255;
                dst[i + 3] = alpha; // smooth edges
            }

            Marshal.Copy(dst, 0, dstData.Scan0, bytes);
            return output;
        }
        finally
        {
            input.UnlockBits(srcData);
            output.UnlockBits(dstData);
        }
    }



}