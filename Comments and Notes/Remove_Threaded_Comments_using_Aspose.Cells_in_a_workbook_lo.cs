using System;
using Aspose.Cells;

namespace RemoveThreadedCommentsDemo
{
    class Program
    {
        static void Main()
        {
            // Load the existing XLSX workbook
            Workbook workbook = new Workbook("InputWorkbook.xlsx");

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Clear all comments (including threaded comments) from the worksheet
                sheet.ClearComments();
            }

            // Save the workbook after removing threaded comments
            workbook.Save("OutputWorkbook.xlsx", SaveFormat.Xlsx);
        }
    }
}