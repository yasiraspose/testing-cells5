using System;
using Aspose.Cells;

namespace MergeWorkbooksDemo
{
    class Program
    {
        static void Main()
        {
            // Create the source workbook and add some data
            Workbook sourceWorkbook = new Workbook();                     // workbook-create
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];
            sourceSheet.Cells["A1"].PutValue("Source Data");             // cell-value

            // Create the destination workbook and add some data
            Workbook destWorkbook = new Workbook();                       // workbook-create
            Worksheet destSheet = destWorkbook.Worksheets[0];
            destSheet.Cells["B2"].PutValue("Destination Data");         // cell-value

            // Combine the source workbook into the destination workbook
            destWorkbook.Combine(sourceWorkbook);                         // Combine method (no specific rule)

            // Save the combined workbook to disk
            string outputPath = "CombinedWorkbook.xlsx";
            destWorkbook.Save(outputPath, SaveFormat.Xlsx);              // workbook-save
        }
    }
}