using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;

class AddTimelineFromTemplate
{
    static void Main()
    {
        // Load the XLTX template workbook
        Workbook workbook = new Workbook("TemplateFile.xltx"); // workbook-load rule

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0]; // worksheet-access rule

        // ------------------------------------------------------------
        // Assume the template contains a date column in A1:A4 with header "Date"
        // Create a pivot table that will serve as the data source for the timeline
        // ------------------------------------------------------------
        int pivotIndex = sheet.PivotTables.Add("A1:A4", "C1", "PivotTable1");
        PivotTable pivot = sheet.PivotTables[pivotIndex];
        pivot.AddFieldToArea(PivotFieldType.Row, "Date");
        pivot.RefreshData();
        pivot.CalculateData();

        // ------------------------------------------------------------
        // Add a timeline linked to the pivot table.
        // The timeline will be placed with its upper‑left corner at cell E1.
        // ------------------------------------------------------------
        int timelineIndex = sheet.Timelines.Add(pivot, "E1", "Date");
        Timeline timeline = sheet.Timelines[timelineIndex];
        timeline.Name = "MyTimeline";

        // Save the modified workbook
        workbook.Save("ResultWithTimeline.xlsx", SaveFormat.Xlsx); // workbook-save rule
    }
}