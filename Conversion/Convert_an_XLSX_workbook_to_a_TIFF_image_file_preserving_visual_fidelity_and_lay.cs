using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

class XlsxToTiffConverter
{
    static void Main()
    {
        // Input Excel file path
        string inputPath = "input.xlsx";

        // Output TIFF file path
        string outputPath = "output.tiff";

        // Load the workbook from the XLSX file (lifecycle rule: workbook-load)
        Workbook workbook = new Workbook(inputPath);

        // Configure image rendering options for high‑quality TIFF output
        ImageOrPrintOptions options = new ImageOrPrintOptions
        {
            ImageType = ImageType.Tiff,                     // Set output format to TIFF
            TiffCompression = TiffCompression.CompressionLZW, // Use lossless LZW compression
            HorizontalResolution = 300,                     // 300 DPI horizontal resolution
            VerticalResolution = 300,                       // 300 DPI vertical resolution
            OnePagePerSheet = true                          // Each worksheet becomes a separate page in the TIFF
        };

        // Create a renderer for the entire workbook
        WorkbookRender renderer = new WorkbookRender(workbook, options);

        // Render the whole workbook to a multi‑page TIFF file (preserves layout and visual fidelity)
        renderer.ToImage(outputPath);

        Console.WriteLine($"Workbook successfully converted to TIFF: {outputPath}");
    }
}