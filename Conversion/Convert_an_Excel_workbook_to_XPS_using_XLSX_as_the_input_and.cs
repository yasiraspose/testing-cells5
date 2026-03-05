using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsConversionDemo
{
    class Program
    {
        static void Main()
        {
            // Define source XLSX file and destination XPS file paths
            string sourcePath = "input.xlsx";
            string destPath   = "output.xps";

            // Convert the Excel workbook to XPS format using Aspose.Cells ConversionUtility
            // This utilizes the provided Convert(string, string) method as required by the rules
            ConversionUtility.Convert(sourcePath, destPath);

            Console.WriteLine("Conversion from XLSX to XPS completed successfully.");
        }
    }
}