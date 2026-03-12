using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;

class InsertTimelineFromTxt
{
    static void Main()
    {
        string txtPath = "data.txt";
        string outputPath = "TimelineFromTxt.xlsx";

        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        string[] lines = File.ReadAllLines(txtPath);
        if (lines.Length == 0) return;

        // Write header
        string[] headers = lines[0].Split('\t');
        for (int c = 0; c < headers.Length; c++)
        {
            sheet.Cells[0, c].PutValue(headers[c]);
        }

        // Write data rows
        for (int r = 1; r < lines.Length; r++)
        {
            string[] parts = lines[r].Split('\t');
            for (int c = 0; c < parts.Length; c++)
            {
                if (c == 0) // first column assumed to be a date
                {
                    if (DateTime.TryParse(parts[c], out DateTime dt))
                        sheet.Cells[r, c].PutValue(dt);
                    else
                        sheet.Cells[r, c].PutValue(parts[c]);
                }
                else
                {
                    if (double.TryParse(parts[c], NumberStyles.Any, CultureInfo.InvariantCulture, out double num))
                        sheet.Cells[r, c].PutValue(num);
                    else
                        sheet.Cells[r, c].PutValue(parts[c]);
                }
            }
        }

        int lastRow = sheet.Cells.MaxDataRow;
        int lastCol = sheet.Cells.MaxDataColumn;
        string sourceRange = $"A1:{CellIndexToName(lastRow, lastCol)}";

        int pivotIdx = sheet.PivotTables.Add(sourceRange, "E1", "PivotTable1");
        PivotTable pivot = sheet.PivotTables[pivotIdx];

        string dateFieldName = sheet.Cells[0, 0].StringValue;
        pivot.AddFieldToArea(PivotFieldType.Row, dateFieldName);

        if (lastCol >= 1)
        {
            string dataFieldName = sheet.Cells[0, 1].StringValue;
            pivot.AddFieldToArea(PivotFieldType.Data, dataFieldName);
        }

        pivot.RefreshData();
        pivot.CalculateData();

        int timelineIdx = sheet.Timelines.Add(pivot, "G1", dateFieldName);
        sheet.Timelines[timelineIdx].Caption = "Date Timeline";

        workbook.Save(outputPath, SaveFormat.Xlsx);
    }

    static string CellIndexToName(int row, int col)
    {
        int dividend = col + 1;
        string columnName = string.Empty;
        while (dividend > 0)
        {
            int modulo = (dividend - 1) % 26;
            columnName = Convert.ToChar('A' + modulo) + columnName;
            dividend = (dividend - modulo) / 26;
        }
        return $"{columnName}{row + 1}";
    }
}