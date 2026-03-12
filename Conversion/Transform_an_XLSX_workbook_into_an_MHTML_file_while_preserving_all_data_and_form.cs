using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the source XLSX file
        string sourcePath = "input.xlsx";

        // Desired path for the MHTML output
        string destinationPath = "output.mht";

        // Load the workbook from the XLSX file (lifecycle rule: workbook-load)
        Workbook workbook = new Workbook(sourcePath);

        // Save the workbook as MHTML, preserving all data and formatting
        // (lifecycle rule: workbook-save)
        workbook.Save(destinationPath, SaveFormat.MHtml);

        Console.WriteLine("Workbook successfully converted to MHTML.");
    }
}