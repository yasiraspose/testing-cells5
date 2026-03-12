using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;

class InsertTimelineWithNumberFormat
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // -------------------------------------------------
        // Populate worksheet with sample data (existing data)
        // -------------------------------------------------
        sheet.Cells["A1"].PutValue("Date");   // Header for date column
        sheet.Cells["B1"].PutValue("Value"); // Header for value column

        sheet.Cells["A2"].PutValue(new DateTime(2023, 1, 1));
        sheet.Cells["B2"].PutValue(120);

        sheet.Cells["A3"].PutValue(new DateTime(2023, 2, 1));
        sheet.Cells["B3"].PutValue(150);

        sheet.Cells["A4"].PutValue(new DateTime(2023, 3, 1));
        sheet.Cells["B4"].PutValue(180);

        // -------------------------------------------------
        // Apply NUMBER (date) format to the date column
        // -------------------------------------------------
        // Retrieve the style from a date cell, modify it, and reapply
        Cell dateCell = sheet.Cells["A2"];
        Style dateStyle = dateCell.GetStyle();
        dateStyle.Custom = "mm/dd/yyyy"; // NUMBERS format for dates
        // Apply the same style to all date cells to keep data integrity
        for (int row = 2; row <= 4; row++)
        {
            Cell c = sheet.Cells[row - 1, 0]; // column A (index 0)
            c.SetStyle(dateStyle);
        }

        // -------------------------------------------------
        // Create a PivotTable that will serve as the Timeline source
        // -------------------------------------------------
        int pivotIdx = sheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
        PivotTable pivot = sheet.PivotTables[pivotIdx];
        pivot.AddFieldToArea(PivotFieldType.Row, "Date");
        pivot.AddFieldToArea(PivotFieldType.Data, "Value");
        pivot.RefreshData();
        pivot.CalculateData();

        // -------------------------------------------------
        // Add a Timeline linked to the PivotTable
        // -------------------------------------------------
        // Place the Timeline starting at cell F1
        int timelineIdx = sheet.Timelines.Add(pivot, "F1", "Date");
        Timeline timeline = sheet.Timelines[timelineIdx];

        // Optional: set a caption for clarity
        timeline.Caption = "Sales Timeline";

        // -------------------------------------------------
        // Save the workbook
        // -------------------------------------------------
        workbook.Save("TimelineNumbersFormat.xlsx", SaveFormat.Xlsx);
    }
}