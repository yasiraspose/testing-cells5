using System;
using Aspose.Cells;

class SaveWorkbookExample
{
    static void Main()
    {
        // Create a new workbook (workbook-create rule)
        Workbook workbook = new Workbook();

        // Access the first worksheet and add sample data
        Worksheet worksheet = workbook.Worksheets[0];
        worksheet.Cells["A1"].PutValue("Hello");
        worksheet.Cells["B1"].PutValue("World");

        // Define the output file path
        string outputPath = "output.xlsx";

        // Save the workbook to XLSX format (workbook-save rule)
        workbook.Save(outputPath, SaveFormat.Xlsx);
    }
}