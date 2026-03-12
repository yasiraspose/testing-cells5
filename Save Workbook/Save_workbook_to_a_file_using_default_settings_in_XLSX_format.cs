using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule: workbook-create)
        Workbook workbook = new Workbook();

        // Optional: add some sample data
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Hello World");

        // Save the workbook to an XLSX file using default settings (lifecycle rule: workbook-save)
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}