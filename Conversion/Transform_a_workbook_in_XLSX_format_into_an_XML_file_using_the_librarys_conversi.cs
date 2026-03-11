using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsConversionDemo
{
    public class XlsxToXmlConverter
    {
        public static void Run()
        {
            string sourcePath = "input.xlsx";
            string destPath = "output.xml";

            // Load the workbook (optional, ensures file is valid)
            Workbook workbook = new Workbook(sourcePath);

            // Convert the Excel file to XML
            ConversionUtility.Convert(sourcePath, destPath);

            Console.WriteLine($"Conversion completed: '{sourcePath}' -> '{destPath}'");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            XlsxToXmlConverter.Run();
        }
    }
}