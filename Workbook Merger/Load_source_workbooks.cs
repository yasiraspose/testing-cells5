using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsLoadExample
{
    class Program
    {
        static void Main(string[] args)
        {
            // Define the paths to the source workbooks
            string sourcePath1 = "SourceWorkbook1.xlsx";
            string sourcePath2 = "SourceWorkbook2.xlsx";

            // Load or create the first workbook
            Workbook workbook1 = LoadOrCreateWorkbook(sourcePath1, "Workbook1 Sheet1 A1");

            // Load or create the second workbook
            Workbook workbook2 = LoadOrCreateWorkbook(sourcePath2, "Workbook2 Sheet1 A1");

            // Access the first worksheet from each workbook
            Worksheet sheet1 = workbook1.Worksheets[0];
            Worksheet sheet2 = workbook2.Worksheets[0];

            // Read and display the value from cell A1 of each worksheet
            Console.WriteLine("Workbook 1, Sheet 1, A1: " + sheet1.Cells["A1"].StringValue);
            Console.WriteLine("Workbook 2, Sheet 1, A1: " + sheet2.Cells["A1"].StringValue);
        }

        private static Workbook LoadOrCreateWorkbook(string path, string a1Value)
        {
            if (File.Exists(path))
            {
                return new Workbook(path);
            }
            else
            {
                // Create a new workbook with a single worksheet and set A1 value
                Workbook wb = new Workbook();
                Worksheet ws = wb.Worksheets[0];
                ws.Name = Path.GetFileNameWithoutExtension(path);
                ws.Cells["A1"].PutValue(a1Value);
                // Save the workbook so subsequent runs can load it
                wb.Save(path);
                return wb;
            }
        }
    }
}