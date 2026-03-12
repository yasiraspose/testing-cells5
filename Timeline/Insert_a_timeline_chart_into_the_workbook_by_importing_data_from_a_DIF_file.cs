using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;

class InsertTimelineFromDif
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data (A1:C10) with headers and values
        sheet.Cells["A1"].PutValue("Date");
        sheet.Cells["B1"].PutValue("Category");
        sheet.Cells["C1"].PutValue("Value");

        for (int i = 2; i <= 10; i++)
        {
            sheet.Cells[$"A{i}"].PutValue(DateTime.Today.AddDays(i - 2));
            sheet.Cells[$"B{i}"].PutValue("CategoryA");
            sheet.Cells[$"C{i}"].PutValue(i * 10);
        }

        // Add a PivotTable based on the sample data
        int pivotIndex = sheet.PivotTables.Add("A1:C10", "E1", "PivotTable1");
        PivotTable pivot = sheet.PivotTables[pivotIndex];

        // Configure PivotTable fields
        pivot.AddFieldToArea(PivotFieldType.Row, "Date");
        pivot.AddFieldToArea(PivotFieldType.Data, "Value");

        // Refresh and calculate the PivotTable
        pivot.RefreshData();
        pivot.CalculateData();

        // Add a Timeline linked to the PivotTable
        int timelineIndex = sheet.Timelines.Add(pivot, "G1", "Date");
        Timeline timeline = sheet.Timelines[timelineIndex];
        timeline.Caption = "Sales Over Time";

        // Save the workbook
        workbook.Save("TimelineResult.xlsx", SaveFormat.Xlsx);
    }
}