using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class MergeChartsDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data (Categories in A, Values for first chart in B, Values for second chart in C)
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("ColumnSeries");
        sheet.Cells["C1"].PutValue("LineSeries");
        string[] categories = { "A", "B", "C", "D", "E" };
        for (int i = 0; i < categories.Length; i++)
        {
            sheet.Cells[i + 1, 0].PutValue(categories[i]);          // Column A
            sheet.Cells[i + 1, 1].PutValue((i + 1) * 10);         // Column B
            sheet.Cells[i + 1, 2].PutValue((i + 1) * 15);         // Column C
        }

        // Add the first chart (Column) – placed near the top of the sheet
        int firstChartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart firstChart = sheet.Charts[firstChartIdx];
        // Set data source for the first chart (uses column B)
        firstChart.NSeries.Add("B2:B6", true);
        firstChart.Title.Text = "Merged Chart";

        // Add line series to the same chart (using column C)
        firstChart.NSeries.Add("C2:C6", true);
        firstChart.NSeries[1].Type = ChartType.Line;

        // Save the workbook
        workbook.Save("MergeChartsDemo.xlsx");
    }
}