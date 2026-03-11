using System;
using Aspose.Cells;

namespace AsposeCellsConversion
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source XLSX file
            string sourcePath = "input.xlsx";

            // Path where the ODS file will be saved
            string destinationPath = "output.ods";

            // Load the existing XLSX workbook (workbook-load rule)
            Workbook workbook = new Workbook(sourcePath);

            // Save the workbook as ODS format (workbook-save rule)
            workbook.Save(destinationPath, SaveFormat.Ods);

            Console.WriteLine($"Conversion completed: '{sourcePath}' -> '{destinationPath}'");
        }
    }
}