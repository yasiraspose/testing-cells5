using System;
using Aspose.Cells;

namespace AsposeCellsThreadedCommentRemoval
{
    class Program
    {
        static void Main()
        {
            // Load the existing XLSX workbook
            Workbook workbook = new Workbook("input.xlsx");

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Determine the used range of the worksheet
                int maxRow = sheet.Cells.MaxDataRow;
                int maxCol = sheet.Cells.MaxDataColumn;

                // Loop through all cells within the used range
                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        // Get the threaded comments for the current cell
                        ThreadedCommentCollection threadedComments = sheet.Comments.GetThreadedComments(row, col);

                        // If there are any threaded comments, remove them all
                        if (threadedComments != null && threadedComments.Count > 0)
                        {
                            // Remove from the end to avoid index shifting
                            for (int i = threadedComments.Count - 1; i >= 0; i--)
                            {
                                threadedComments.RemoveAt(i);
                            }
                        }
                    }
                }
            }

            // Save the workbook after removing all threaded comments
            workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
    }
}