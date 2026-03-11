using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create the first workbook and add some data
        var sourceWorkbook = new Workbook();
        sourceWorkbook.Worksheets[0].Cells["A1"].PutValue("Data from Source Workbook");

        // Create the second workbook (destination) and add some data
        var destWorkbook = new Workbook();
        destWorkbook.Worksheets[0].Cells["B2"].PutValue("Data from Destination Workbook");

        // Merge the source workbook into the destination workbook
        destWorkbook.Combine(sourceWorkbook);

        // Save the combined workbook
        destWorkbook.Save("CombinedWorkbook.xlsx", SaveFormat.Xlsx);
    }
}