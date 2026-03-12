using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;
using Aspose.Cells.Rendering;

class InsertTimelineWithHtml
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook(); // workbook-create

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0]; // worksheet-access

        // Populate sample data (Date and Sales)
        sheet.Cells["A1"].PutValue("Date");
        sheet.Cells["B1"].PutValue("Sales");
        sheet.Cells["A2"].PutValue(new DateTime(2023, 1, 1));
        sheet.Cells["A3"].PutValue(new DateTime(2023, 2, 1));
        sheet.Cells["A4"].PutValue(new DateTime(2023, 3, 1));
        sheet.Cells["B2"].PutValue(1000);
        sheet.Cells["B3"].PutValue(2000);
        sheet.Cells["B4"].PutValue(3000);

        // Add a PivotTable that will serve as the data source for the Timeline
        int pivotIdx = sheet.PivotTables.Add("A1:B4", "D1", "SalesPivot");
        PivotTable pivot = sheet.PivotTables[pivotIdx];
        pivot.AddFieldToArea(PivotFieldType.Row, "Date");
        pivot.AddFieldToArea(PivotFieldType.Data, "Sales");
        pivot.RefreshData();
        pivot.CalculateData();

        // Add a Timeline linked to the PivotTable
        // Upper‑left corner of the Timeline will start at cell A6 (row 5, column 0)
        int timelineIdx = sheet.Timelines.Add(pivot, 5, 0, "Date");
        Timeline timeline = sheet.Timelines[timelineIdx];

        // Optional: set some visual properties via the underlying shape
        timeline.Shape.Top = 150;      // vertical offset in pixels
        timeline.Shape.Left = 20;      // horizontal offset in pixels
        timeline.Shape.Width = 400;    // width in pixels
        timeline.Shape.Height = 80;    // height in pixels
        timeline.Caption = "Sales Timeline";

        // Insert HTML‑formatted text into a cell to demonstrate cross‑platform rendering
        // The HTML will be preserved when saving to HTML because ParseHtmlTagInCell is true
        sheet.Cells["F1"].PutValue("<b>Important:</b> Use the timeline above to filter sales data.");

        // Configure HTML save options
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.ParseHtmlTagInCell = true;          // preserve HTML tags in cell values
        htmlOptions.ExportImagesAsBase64 = true;       // embed images (including the timeline) as Base64
        htmlOptions.IsExportComments = false;          // no comments needed
        htmlOptions.ExportGridLines = true;            // show grid lines for better readability

        // Save the workbook as an HTML file
        workbook.Save("TimelineWithHtml.html", htmlOptions); // workbook-save
    }
}