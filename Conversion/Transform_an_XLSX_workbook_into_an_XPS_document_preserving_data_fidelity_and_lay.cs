using System;
using Aspose.Cells;

namespace AsposeCellsXpsConversion
{
    class Program
    {
        static void Main(string[] args)
        {
            // Input XLSX file path
            string inputPath = "input.xlsx";

            // Output XPS file path
            string outputPath = "output.xps";

            // Load the workbook (lifecycle rule: workbook-load)
            Workbook workbook = new Workbook(inputPath);

            // Save the workbook as XPS (lifecycle rule: workbook-save with SaveFormat)
            workbook.Save(outputPath, SaveFormat.Xps);

            Console.WriteLine($"Conversion completed: '{inputPath}' -> '{outputPath}'");
        }
    }
}