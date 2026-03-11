using System;
using Aspose.Cells;

class ConvertXlsxToMhtml
{
    static void Main()
    {
        // Path to the source XLSX workbook
        string inputPath = "input.xlsx";

        // Desired output MHTML file path
        string outputPath = "output.mht";

        // Load the existing workbook (uses workbook-load rule)
        Workbook workbook = new Workbook(inputPath);

        // Save the workbook as MHTML, preserving all data and formatting (uses workbook-save rule)
        workbook.Save(outputPath, SaveFormat.MHtml);

        Console.WriteLine($"Workbook successfully converted to MHTML: {outputPath}");
    }
}