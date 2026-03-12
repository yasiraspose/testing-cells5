using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;

class TimelineFromTabDelimited
{
    static void Main()
    {
        // Paths for the input TAB‑delimited file and the output workbook
        string inputPath = "data.tsv";          // <-- replace with your actual file path
        string outputPath = "TimelineResult.xlsx";

        // Create a new workbook (lifecycle rule: workbook-create)
        Workbook workbook = new Workbook();

        // Access the first worksheet (lifecycle rule: worksheet-access)
        Worksheet sheet = workbook.Worksheets[0];

        // Import the TAB‑delimited text file into the worksheet starting at cell A1
        // (lifecycle rule: ImportCSV with splitter "\t")
        sheet.Cells.ImportCSV(inputPath, "\t", true, 0, 0);

        // ------------------------------------------------------------
        // Create a PivotTable based on the imported data.
        // Assume the data occupies columns A‑C and rows 1‑100.
        // Adjust the range as needed for your actual data.
        // ------------------------------------------------------------
        string sourceRange = "A1:C100";          // source data range
        string pivotDest   = "E1";               // where the pivot table will be placed
        int pivotIndex = sheet.PivotTables.Add(sourceRange, pivotDest, "PivotTable1");
        PivotTable pivot = sheet.PivotTables[pivotIndex];

        // Add fields to the PivotTable.
        // Replace "Date", "Category", "Value" with the actual column headers in your file.
        pivot.AddFieldToArea(PivotFieldType.Row, "Date");
        pivot.AddFieldToArea(PivotFieldType.Column, "Category");
        pivot.AddFieldToArea(PivotFieldType.Data, "Value");

        // Refresh the PivotTable so it reflects the imported data.
        pivot.RefreshData();
        pivot.CalculateData();

        // ------------------------------------------------------------
        // Add a Timeline control linked to the PivotTable.
        // The Timeline will be placed with its upper‑left corner at cell G1.
        // ------------------------------------------------------------
        sheet.Timelines.Add(pivot, "G1", "Date");

        // Save the workbook (lifecycle rule: workbook-save)
        workbook.Save(outputPath, SaveFormat.Xlsx);
    }
}