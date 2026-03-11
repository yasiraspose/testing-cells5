using System;
using Aspose.Cells;

namespace CombineWorkbooksExample
{
    class Program
    {
        static void Main()
        {
            // Create the source workbook and add some data
            Workbook sourceWorkbook = new Workbook();                     // workbook-create
            sourceWorkbook.Worksheets[0].Cells["A1"].PutValue("Source Data"); // cell-value

            // Create the destination workbook and add some data
            Workbook destWorkbook = new Workbook();                       // workbook-create
            destWorkbook.Worksheets[0].Cells["B2"].PutValue("Destination Data"); // cell-value

            // Combine the source workbook into the destination workbook
            destWorkbook.Combine(sourceWorkbook); // Workbook.Combine method

            // Save the combined workbook as XLSX
            destWorkbook.Save("CombinedWorkbook.xlsx", SaveFormat.Xlsx); // workbook-save
        }
    }
}