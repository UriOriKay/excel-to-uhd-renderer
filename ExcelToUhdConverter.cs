using PdfiumViewer;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

internal sealed class ExcelToUhdConverter
{
    public void Convert(string inputPath, string pngOutputDir, Action<string>? logger = null)
    {
        // Create a unique temporary PDF path (intermediate file)
        string tempPdfPath = Path.Combine(
            Path.GetTempPath(),
            $"{Path.GetFileNameWithoutExtension(inputPath)}-{Guid.NewGuid():N}.pdf"
            );

        Convert(inputPath, tempPdfPath, pngOutputDir, logger);
    }

    public void Convert(string inputPath, string outputPdfPath, string pngOutputDir, Action<string>? logger = null)
    {
        try
        {

            string baseName = Path.GetFileNameWithoutExtension(inputPath);

            ConvertExcelToPdf(inputPath, outputPdfPath);
            ConvertPdfToPngPages(outputPdfPath, pngOutputDir, baseName, dpi: 400);
        }
        finally
        {
            DeletePdfFile(outputPdfPath,logger);
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

            // Safety check: ensure PDF was actually created
            if (!File.Exists(outputPath))
                throw new IOException("Excel export reported success, but PDF was not created: " + outputPath);
        }
        finally
        {
            // CLose workbook safely
            if (workbook != null)
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
    private static void ConvertPdfToPngPages(string pdfPath, string outputDir, string baseName, int dpi = 300)
    {
        if (!File.Exists(pdfPath))
            throw new FileNotFoundException("PDF not found", pdfPath);

        Directory.CreateDirectory(outputDir);

        using var doc = PdfDocument.Load(pdfPath);

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

    /// <summary>
    /// Converts a rendered PDF/Excel page bitmap (typically black text on white background)
    /// into a new bitmap containing white text/lines on a transparent background.
    /// </summary>
    /// <param name="input">
    /// Source bitmap to process. Expected to be <see cref="PixelFormat.Format32bppArgb"/>.
    /// Dark pxles are treated as foreground (text/lines), bright pixels as background.
    /// </param>
    /// <param name="backgroundCutoff">
    /// Brightness threshold (0..255) at or above which pixels are treated as background and become fully transparent.
    /// Higher values keep more near-white pixels transparent (good for clean white backgrounds).
    /// Default is 245. 
    /// </param>
    /// <param name="gamma">
    /// Gamma adjustment applied to the computed alpha of foreground pixels.
    /// Use values &lt; 1.0 to make edges stronger/inkier, and values &gt; 1.0 to make edges thinner/softer.
    /// Default is 1.0 (no adjustment).
    /// </param>
    /// <returns>
    /// A new <see cref="Bitmap"/> in <see cref="PixelFormat.Format32bppArgb"/> where:
    /// - RGB is always white (255,255,255)
    /// - Alpha encodes the original pixel darkness (anti-aliased edges become partially transparent)
    /// </returns>
    /// <remarks>
    /// Implementation details:
    /// - Computes pixel luminance using Rec.709 coefficients (0.2126 R + 0.7152 G + 0.0722 B).
    /// - Background pixels (lum &gt;= <paramref name="backgroundCutoff"/>) become fully transparent.
    /// - Foreground pixels become white with alpha derived from "darkness": alpha = 255 - lum.
    /// - The alpha can be gamma-shaped to control perceived stroke weight.
    /// 
    /// Typical use case:
    /// - After rasterizing a PDF page to an image, call this method to obtain a white-on-transparent overlay.
    /// - Draw that overlay onto a black UHD canvas to produce "white text on black background" output.
    /// </remarks>
    /// <exception cref="ArgumentNullException">
    /// Thrown if <paramref name="input"/> is null.
    /// </exception>
    /// <exception cref="ArgumentOutOfRangeException">
    /// Thrown if <paramref name="backgroundCutoff"/> or <paramref name="gamma"/> are outside valid ranges.
    /// </exception>
    /// <exception cref="ArgumentException">
    /// Thrown if <paramref name="input"/> is not in <see cref="PixelFormat.Format32bppArgb"/>.
    /// </exception>
    private static Bitmap ToWhiteTextTransparentBackground(Bitmap input, byte backgroundCutoff = 245, float gamma = 1.0f)
    {
        if (input is null) throw new ArgumentNullException(nameof(input));
        if (gamma <= 0f) throw new ArgumentOutOfRangeException(nameof(gamma), "Gamma must be > 0.");
        if (input.PixelFormat != PixelFormat.Format32bppArgb)
            throw new ArgumentException("Input bitmap must be Format32bppArgb.", nameof(input));

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

    /// <summary>
    /// Creates a UHD Image (default 3840x2160) with a solid black background
    /// and draws the provided overlay image centered and proportionally scaled
    /// onto the canvas without distortion.
    /// </summary>
    /// <param name="overlay">
    /// The source image that will be drawn onto the UHD canvas.
    /// The image is scaled proportionally to fit within the target dimensions
    /// while preserving its aspect ratio.
    /// </param>
    /// <param name="outputFile">
    /// Full file path where the resulting PNG image will be written.
    /// Existing files will be overwritten.
    /// </param>
    /// <param name="targetW">
    /// Taeget canvas width in pixels. Default is 3840 (UHD width)
    /// </param>
    /// <param name="targetH">
    /// Target canvas height in pixels. Default is 2160 (UHD height). 
    /// </param>
    /// <remarks>
    /// Rendering behavior:
    /// - The background is filled with solid black.
    /// - The overlay image is scaled using aspect-ratio-preserving fit logic.
    /// - No stretching or distortion is applied.
    /// - If aspect ratios differ, letterboxing (black margins) will appear.
    /// - High quality bicubic interpolation is used for scaling.
    /// 
    /// This method assumes the overlay image may contain transparency
    /// (e.g. white text with alpha channel).
    /// </remarks>
    /// <exception cref="ArgumentNullException">
    /// Thrown if the overlay image is null.
    /// </exception>
    /// <exception cref="ExternalException">
    /// Thrown if saving the PNG file fails.
    /// </exception>
    private static void SaveUhdCanvasWithCenteredImage(Image overlay, string outputFile, int targetW = 3840, int targetH = 2160)
    {
        if (overlay == null)
            throw new ArgumentNullException(nameof(overlay));

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

    /// <summary>
    /// Attempts to delete the specified file.
    /// </summary>
    /// <param name="path">
    /// The full path of the file to delete
    /// </param>
    /// <param name="logger">
    /// Optional logging callback invoked if deletion fails.
    /// </param>
    /// <remarks>
    /// This method performs a best-effort deletion.
    /// Exceptions are not rethrown but can be logged via provided logger.
    /// </remarks>
    private static void DeletePdfFile(string path, Action<string>? logger = null)
    {
        try
        {
            if (File.Exists(path))
                File.Delete(path);
        }
        catch (Exception ex)
        {
            logger?.Invoke($"Warning: failed to delete temporary file '{path}'. {ex.Message} ");
        }
    }
}