using System;
using System.Text;
using Aspose.Cells;

class XlsxToCsvConverter
{
    static void Main()
    {
        // Path to the source XLSX file
        string inputPath = "input.xlsx";

        // Path for the resulting CSV file
        string outputPath = "output.csv";

        // Load the workbook (uses workbook-load rule)
        Workbook workbook = new Workbook(inputPath);

        // Configure save options to trim leading empty rows and columns
        TxtSaveOptions saveOptions = new TxtSaveOptions
        {
            TrimLeadingBlankRowAndColumn = true,
            Encoding = Encoding.UTF8
        };

        // Save the workbook as CSV with the specified options (uses workbook-save rule)
        workbook.Save(outputPath, saveOptions);

        Console.WriteLine("XLSX has been converted to CSV with leading blanks trimmed.");
    }
}