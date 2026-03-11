using System;
using Aspose.Cells;

namespace MergeSpecificWorksheets
{
    class Program
    {
        static void Main()
        {
            // Create the destination workbook (empty)
            Workbook destWorkbook = new Workbook();

            // Files to merge and the worksheet names to extract
            string[] sourceFiles = { "Source1.xlsx", "Source2.xlsx" };
            string[] sheetNamesToMerge = { "Data1", "Data2" };

            foreach (string filePath in sourceFiles)
            {
                // Load each source workbook
                Workbook srcWorkbook = new Workbook(filePath);

                foreach (string sheetName in sheetNamesToMerge)
                {
                    // Access the specific worksheet by name
                    Worksheet srcSheet = srcWorkbook.Worksheets[sheetName];
                    if (srcSheet == null)
                        continue; // Worksheet not found; skip to next

                    // Add a new worksheet to the destination workbook with a unique name
                    string newSheetName = $"{srcSheet.Name}_{System.IO.Path.GetFileNameWithoutExtension(filePath)}";
                    Worksheet destSheet = destWorkbook.Worksheets.Add(newSheetName);

                    // Copy the contents and formats from the source sheet
                    destSheet.Copy(srcSheet);
                }
            }

            // Save the merged workbook
            destWorkbook.Save("MergedWorkbook.xlsx", SaveFormat.Xlsx);
        }
    }
}