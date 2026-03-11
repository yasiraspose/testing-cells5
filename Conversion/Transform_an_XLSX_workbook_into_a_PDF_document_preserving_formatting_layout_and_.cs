using System;
using Aspose.Cells;               // Core Aspose.Cells namespace
using Aspose.Cells.Utility;      // For ConversionUtility (optional)

class XlsxToPdfConverter
{
    static void Main()
    {
        // Path to the source XLSX file
        string sourcePath = "input.xlsx";

        // Desired PDF output path
        string pdfPath = "output.pdf";

        // Load the workbook from the XLSX file (workbook-load rule)
        Workbook workbook = new Workbook(sourcePath);

        // Save the workbook as PDF, preserving formatting, layout and pagination (workbook-save rule)
        workbook.Save(pdfPath, SaveFormat.Pdf);

        Console.WriteLine("Conversion completed: " + pdfPath);
    }
}