# excel-to-uhd-renderer

A Windows-based CLI tool that converts Excel files into UHD (3840x2160) PNG images with a black background and white text.

This project demonstrates:
- COM interop with Microsoft Excel
- PDF generation via Excel
- PDF rasterization using PdfiumViewer
- Custom image processing (alpha-based text extraction)
- High-quality UHD canvas rendering

---

## Features

- Convert `.xlsx` → `.pdf`
- Render each PDF page → PNG
- Generate UHD (3840x2160) images
- Black background with white text
- Anti-aliased text preserved via alpha-based luminance mapping
- Aspect-ratio safe scaling (no distortion)

---

## Rendering Pipeline

Excel (.xlsx)  
→ Export as PDF  
→ Rasterize each PDF page  
→ Convert black text on white background  
→ Extract white text with alpha transparency  
→ Draw centered onto UHD canvas (3840x2160)  
→ Save as PNG  

---

## 📦 Requirements

- Windows OS
- .NET 10 (or compatible .NET version)
- Microsoft Excel installed
- x64 build configuration

NuGet dependencies:
- PdfiumViewer
- PdfiumViewer.Native.x86_64.v8-xfa
- System.Drawing.Common

---

## 🖥️ Usage

```bash
dotnet run -- input.xlsx output.pdf ./output-folder