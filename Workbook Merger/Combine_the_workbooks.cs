using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create the source workbook and add some data
        Workbook sourceWorkbook = new Workbook();                     // workbook-create
        sourceWorkbook.Worksheets[0].Cells["A1"].PutValue("Source Data"); // cell-value

        // Create the destination workbook and add some data
        Workbook destinationWorkbook = new Workbook();                // workbook-create
        destinationWorkbook.Worksheets[0].Cells["B2"].PutValue("Destination Data"); // cell-value

        // Combine the source workbook into the destination workbook
        destinationWorkbook.Combine(sourceWorkbook);                  // Workbook.Combine

        // Save the combined workbook
        destinationWorkbook.Save("CombinedWorkbook.xlsx", SaveFormat.Xlsx); // workbook-save
    }
}