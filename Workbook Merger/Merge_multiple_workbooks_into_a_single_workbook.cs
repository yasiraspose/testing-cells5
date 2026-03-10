using System;
using System.IO;
using Aspose.Cells;

class MergeWorkbooksDemo
{
    static void Main()
    {
        // List of workbook files to be merged
        string[] sourceFiles = { "File1.xlsx", "File2.xlsx", "File3.xlsx" };
        // Path for the resulting merged workbook
        string outputFile = "MergedWorkbook.xlsx";

        // Ensure source workbooks exist (create dummy workbooks if they don't)
        foreach (string file in sourceFiles)
        {
            if (!File.Exists(file))
            {
                Workbook tempWb = new Workbook();
                tempWb.Worksheets[0].Name = Path.GetFileNameWithoutExtension(file);
                tempWb.Save(file, SaveFormat.Xlsx);
            }
        }

        // Create an empty destination workbook
        Workbook destWorkbook = new Workbook();

        // Load each source workbook and combine it into the destination
        foreach (string file in sourceFiles)
        {
            // Load a source workbook from file
            Workbook srcWorkbook = new Workbook(file);

            // Merge the source workbook into the destination workbook
            destWorkbook.Combine(srcWorkbook);
        }

        // Save the combined workbook to disk
        destWorkbook.Save(outputFile, SaveFormat.Xlsx);
    }
}