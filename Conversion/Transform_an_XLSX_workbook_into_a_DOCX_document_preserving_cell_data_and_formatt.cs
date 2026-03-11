using System;
using Aspose.Cells;

namespace AsposeCellsConversion
{
    class XlsxToDocx
    {
        static void Main()
        {
            // Path to the source XLSX file
            string sourcePath = "input.xlsx";

            // Path for the resulting DOCX file
            string destPath = "output.docx";

            // Load the workbook (uses workbook-load rule)
            Workbook workbook = new Workbook(sourcePath);

            // Save the workbook as DOCX (uses workbook-save rule with SaveFormat.Docx)
            workbook.Save(destPath, SaveFormat.Docx);

            Console.WriteLine("Conversion completed successfully.");
        }
    }
}