using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

class SaveWorkbookDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook(); // workbook-create

        // Access the first worksheet and add sample data
        Worksheet ws = workbook.Worksheets[0]; // worksheet-access
        ws.Cells["A1"].PutValue("Hello");
        ws.Cells["B1"].PutValue("World");

        // 1. Save directly to XLSX file using the (string, SaveFormat) overload
        workbook.Save("output1.xlsx", SaveFormat.Xlsx); // workbook-save

        // 2. Save to a MemoryStream and then write the stream to a file
        using (MemoryStream ms = new MemoryStream())
        {
            workbook.Save(ms, SaveFormat.Xlsx); // free‑form (no specific rule for streams)
            ms.Position = 0;
            using (FileStream file = new FileStream("output2.xlsx", FileMode.Create, FileAccess.Write))
            {
                ms.CopyTo(file);
            }
        }

        // 3. Save using OoxmlSaveOptions to demonstrate custom save options
        OoxmlSaveOptions options = new OoxmlSaveOptions();
        options.CompressionType = OoxmlCompressionType.Level6;
        workbook.Save("output3.xlsx", options); // free‑form (options overload)

        // 4. Convert an existing XLSX file to another XLSX file using ConversionUtility
        // (Shows an alternative saving/conversion method)
        string sourcePath = "output1.xlsx";
        string destPath = "converted.xlsx";
        ConversionUtility.Convert(sourcePath, destPath); // uses Convert method

        Console.WriteLine("All save operations completed successfully.");
    }
}