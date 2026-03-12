using System;
using Aspose.Cells;

class ExcelToPdfConverter
{
    static void Main()
    {
        // Path to the source XLSX file
        string inputPath = "input.xlsx";

        // Desired path for the resulting PDF file
        string outputPath = "output.pdf";

        // Load the workbook from the XLSX file (workbook-load rule)
        Workbook workbook = new Workbook(inputPath);

        // Save the workbook as PDF (workbook-save rule with SaveFormat)
        workbook.Save(outputPath, SaveFormat.Pdf);

        Console.WriteLine("Excel file has been successfully converted to PDF.");
    }
}