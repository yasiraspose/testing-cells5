using System;
using Aspose.Cells;

namespace AsposeCellsMergeExample
{
    class Program
    {
        static void Main()
        {
            // Create the source workbook and add some data
            Workbook sourceWorkbook = new Workbook();                     // workbook-create
            sourceWorkbook.Worksheets[0].Cells["A1"].PutValue("Source Data");
            string sourcePath = "Source.xlsx";
            sourceWorkbook.Save(sourcePath, SaveFormat.Xlsx);            // workbook-save

            // Create the destination workbook and add some data
            Workbook destWorkbook = new Workbook();                       // workbook-create
            destWorkbook.Worksheets[0].Cells["B2"].PutValue("Destination Data");

            // Combine the source workbook into the destination workbook
            destWorkbook.Combine(sourceWorkbook);                         // Combine method (no specific rule)

            // Save the combined workbook
            string combinedPath = "CombinedWorkbook.xlsx";
            destWorkbook.Save(combinedPath, SaveFormat.Xlsx);            // workbook-save

            Console.WriteLine("Workbooks merged successfully.");
        }
    }
}