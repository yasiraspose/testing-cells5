using System;
using Aspose.Cells;

namespace WorkbookMergeExample
{
    class Program
    {
        static void Main()
        {
            // Paths to the workbooks to be merged
            string sourcePath1 = "Source1.xlsx";
            string sourcePath2 = "Source2.xlsx";
            string destinationPath = "CombinedWorkbook.xlsx";

            // Create the destination workbook (empty workbook) - workbook-create rule
            Workbook destinationWorkbook = new Workbook();

            // Load the first source workbook - workbook-load rule
            Workbook sourceWorkbook1 = new Workbook(sourcePath1);

            // Load the second source workbook - workbook-load rule
            Workbook sourceWorkbook2 = new Workbook(sourcePath2);

            // Combine the first source workbook into the destination workbook
            destinationWorkbook.Combine(sourceWorkbook1);

            // Combine the second source workbook into the destination workbook
            destinationWorkbook.Combine(sourceWorkbook2);

            // Save the combined workbook as XLSX - workbook-save rule
            destinationWorkbook.Save(destinationPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbooks merged successfully into '{destinationPath}'.");
        }
    }
}