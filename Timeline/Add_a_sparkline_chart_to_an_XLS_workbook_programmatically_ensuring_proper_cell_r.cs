using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the sparkline (row 0, columns A-D)
        sheet.Cells["A1"].PutValue(5);
        sheet.Cells["B1"].PutValue(2);
        sheet.Cells["C1"].PutValue(1);
        sheet.Cells["D1"].PutValue(3);

        // Define the cell area where the sparkline will be placed (column E, same row)
        CellArea location = new CellArea
        {
            StartRow = 0,
            EndRow = 0,
            StartColumn = 4,
            EndColumn = 4
        };

        // Add a sparkline group of type Line with the data range A1:D1
        int groupIdx = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, location);
        SparklineGroup group = sheet.SparklineGroups[groupIdx];

        // Optional: customize the sparkline group
        group.ShowHighPoint = true;
        group.ShowLowPoint = true;

        // Retrieve the created sparkline (first in the collection)
        Sparkline spark = group.Sparklines[0];

        // Render the sparkline to an image file
        ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
        {
            ImageType = ImageType.Png,
            HorizontalResolution = 300,
            VerticalResolution = 300
        };
        spark.ToImage("sparkline.png", imgOptions);

        // Save the workbook as an XLSX file
        workbook.Save("SparklineDemo.xlsx", SaveFormat.Xlsx);
    }
}