using System;
using Aspose.Cells;

class RemoveThreadedComments
{
    static void Main()
    {
        // Load the existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Remove all comments (including threaded comments) from each worksheet
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            sheet.ClearComments(); // Clears both regular and threaded comments
        }

        // Save the workbook after removing comments
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}