using System;
using System.IO;
using Aspose.Cells;

public class SaveWorkbookToStreamExample
{
    public static void Main()
    {
        Run();
    }

    public static void Run()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Add some sample data
        sheet.Cells["A1"].PutValue("Name");
        sheet.Cells["B1"].PutValue("Score");
        sheet.Cells["A2"].PutValue("Alice");
        sheet.Cells["B2"].PutValue(85);
        sheet.Cells["A3"].PutValue("Bob");
        sheet.Cells["B3"].PutValue(92);

        // Save the workbook to a memory stream in XLSX format
        using (MemoryStream stream = new MemoryStream())
        {
            workbook.Save(stream, SaveFormat.Xlsx);
            stream.Position = 0;

            // Write the stream content to a file for verification
            using (FileStream file = new FileStream("SavedWorkbook.xlsx", FileMode.Create, FileAccess.Write))
            {
                stream.CopyTo(file);
            }

            Console.WriteLine("Workbook successfully saved to stream and written to 'SavedWorkbook.xlsx'.");
        }

        // Dispose the workbook when done
        workbook.Dispose();
    }
}