using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsMergeDemo
{
    class Program
    {
        static void Main()
        {
            // Define temporary files to be merged
            string[] filesToMerge = new string[2];
            filesToMerge[0] = "File1.xlsx";
            filesToMerge[1] = "File2.xlsx";

            // ---------- Create first workbook ----------
            // workbook-create
            Workbook workbook1 = new Workbook();
            // cell-value
            workbook1.Worksheets[0].Cells["A1"].PutValue("File 1 Content");
            // workbook-save
            workbook1.Save(filesToMerge[0]);

            // ---------- Create second workbook ----------
            Workbook workbook2 = new Workbook();
            workbook2.Worksheets[0].Cells["A1"].PutValue("File 2 Content");
            workbook2.Save(filesToMerge[1]);

            // Define cache and output files
            string cachedFile = "CacheFile.tmp";
            string outputFile = "MergedOutput.xlsx";

            try
            {
                // Merge the files using CellsHelper
                CellsHelper.MergeFiles(filesToMerge, cachedFile, outputFile);
                Console.WriteLine($"Files merged successfully. Output saved to: {outputFile}");

                // Load the merged workbook to verify content
                // workbook-load
                Workbook mergedWorkbook = new Workbook(outputFile);

                // Display merged content from the first two worksheets
                Console.WriteLine("Merged content:");
                Console.WriteLine($"Worksheet 0, A1: {mergedWorkbook.Worksheets[0].Cells["A1"].StringValue}");
                Console.WriteLine($"Worksheet 1, A1: {mergedWorkbook.Worksheets[1].Cells["A1"].StringValue}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error merging files: {ex.Message}");
            }
            finally
            {
                // Clean up temporary files
                foreach (var file in filesToMerge)
                {
                    if (File.Exists(file)) File.Delete(file);
                }
                if (File.Exists(cachedFile)) File.Delete(cachedFile);
                // Optionally keep the merged output file; delete if not needed
                // if (File.Exists(outputFile)) File.Delete(outputFile);
            }
        }
    }
}