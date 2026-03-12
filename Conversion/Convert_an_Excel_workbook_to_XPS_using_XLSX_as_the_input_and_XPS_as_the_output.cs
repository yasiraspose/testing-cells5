using System;
using Aspose.Cells;

class ConvertXlsxToXps
{
    static void Main()
    {
        // Path to the source XLSX file
        string inputPath = "input.xlsx";

        // Desired path for the output XPS file
        string outputPath = "output.xps";

        // Load the workbook from the XLSX file
        Workbook workbook = new Workbook(inputPath);

        // Save the workbook in XPS format
        workbook.Save(outputPath, SaveFormat.Xps);
    }
}