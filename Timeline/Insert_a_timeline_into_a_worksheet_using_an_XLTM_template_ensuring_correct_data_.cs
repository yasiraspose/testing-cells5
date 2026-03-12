using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;

class InsertTimelineFromTemplate
{
    static void Main()
    {
        string templatePath = "Template.xltm";
        string outputPath = "WorkbookWithTimeline.xlsx";

        Workbook workbook = new Workbook(templatePath);
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        cells["A1"].PutValue("Date");
        cells["B1"].PutValue("Sales");

        cells["A2"].PutValue(new DateTime(2023, 1, 1));
        cells["B2"].PutValue(1200);
        cells["A3"].PutValue(new DateTime(2023, 2, 1));
        cells["B3"].PutValue(1500);
        cells["A4"].PutValue(new DateTime(2023, 3, 1));
        cells["B4"].PutValue(1800);
        cells["A5"].PutValue(new DateTime(2023, 4, 1));
        cells["B5"].PutValue(2100);

        int pivotIndex = sheet.PivotTables.Add("A1:B5", "D1", "SalesPivot");
        PivotTable pivot = sheet.PivotTables[pivotIndex];
        pivot.AddFieldToArea(PivotFieldType.Row, 0);
        pivot.AddFieldToArea(PivotFieldType.Data, 1);
        pivot.RefreshData();
        pivot.CalculateData();

        // Add Timeline at cell F1 (row 0, column 5)
        int timelineIndex = sheet.Timelines.Add(pivot, 0, 5, "Date");
        Timeline timeline = sheet.Timelines[timelineIndex];
        timeline.Name = "SalesTimeline";
        timeline.Caption = "Sales Over Time";

        workbook.Save(outputPath, SaveFormat.Xlsx);
    }
}