using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class SparklineToXlsm
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule)
        Workbook workbook = new Workbook();

        // Access the first worksheet (worksheet-access rule)
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the sparkline (cell-value rule)
        sheet.Cells["A1"].PutValue(5);
        sheet.Cells["B1"].PutValue(2);
        sheet.Cells["C1"].PutValue(1);
        sheet.Cells["D1"].PutValue(3);

        // Define the cell area where the sparkline will be placed
        CellArea sparklineLocation = new CellArea
        {
            StartRow = 0,   // Row 1 (zero‑based)
            EndRow = 0,
            StartColumn = 4, // Column E (zero‑based)
            EndColumn = 4
        };

        // Add a sparkline group to the worksheet (no specific rule, free‑form)
        int groupIdx = sheet.SparklineGroups.Add(
            SparklineType.Line,          // Sparkline type
            "A1:D1",                     // Data range
            false,                       // Plot by row (horizontal)
            sparklineLocation);          // Location of the sparkline

        SparklineGroup group = sheet.SparklineGroups[groupIdx];

        // Add a sparkline to the group (free‑form)
        // Row 0, Column 4 corresponds to cell E1
        group.Sparklines.Add("A1:D1", 0, 4);

        // Optional: customize appearance (e.g., series color)
        CellsColor seriesColor = workbook.CreateCellsColor();
        seriesColor.Color = System.Drawing.Color.Orange;
        group.SeriesColor = seriesColor;

        // Save the workbook as a macro‑enabled file (workbook-save rule)
        workbook.Save("SparklineExample.xlsm", SaveFormat.Xlsm);
    }
}