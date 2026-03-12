using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

class SparklineDemo
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

        // Define the location where the sparkline will be placed
        CellArea sparklineLocation = new CellArea
        {
            StartRow = 0,
            EndRow = 0,
            StartColumn = 4,
            EndColumn = 4
        };

        // Add a sparkline group with the data range and location (sparkline-groupcollection.Add rule)
        int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, sparklineLocation);
        SparklineGroup group = sheet.SparklineGroups[groupIndex];

        // Add a sparkline to the group (sparklinecollection.Add rule)
        // Row = 0 (first row), Column = 4 (column E)
        int sparklineIndex = group.Sparklines.Add("A1:D1", 0, 4);
        Sparkline sparkline = group.Sparklines[sparklineIndex];

        // Configure visual properties of the sparkline group
        // Series color
        CellsColor seriesColor = workbook.CreateCellsColor();
        seriesColor.Color = Color.Orange;
        group.SeriesColor = seriesColor;

        // High and low point colors
        CellsColor highColor = workbook.CreateCellsColor();
        highColor.Color = Color.Green;
        group.HighPointColor = highColor;
        group.ShowHighPoint = true;

        CellsColor lowColor = workbook.CreateCellsColor();
        lowColor.Color = Color.Red;
        group.LowPointColor = lowColor;
        group.ShowLowPoint = true;

        // Line weight
        group.LineWeight = 1.0;

        // Preset style
        group.PresetStyle = SparklinePresetStyleType.Style5;

        // Optional: display markers and first/last points
        group.ShowMarkers = true;
        group.ShowFirstPoint = true;
        group.ShowLastPoint = true;

        // Save the workbook (lifecycle rule)
        workbook.Save("SparklineDemo.xlsx", SaveFormat.Xlsx);
    }
}