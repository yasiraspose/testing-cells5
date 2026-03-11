using System;
using Aspose.Cells;

namespace AsposeCellsMergeExample
{
    public class WorkbookMerger
    {
        /// <summary>
        /// Merges multiple Excel workbooks into a single workbook.
        /// </summary>
        /// <param name="inputFiles">Array of file paths to be merged.</param>
        /// <param name="outputFile">Path of the resulting merged workbook.</param>
        public static void MergeWorkbooks(string[] inputFiles, string outputFile)
        {
            // Create the destination workbook (empty) – uses workbook-create rule
            Workbook destinationWorkbook = new Workbook();

            // Iterate through each source file, load it, and combine it with the destination
            foreach (string filePath in inputFiles)
            {
                // Load a source workbook – uses workbook-load rule
                Workbook sourceWorkbook = new Workbook(filePath);

                // Combine the source workbook into the destination workbook
                // (No specific rule exists for Combine, so we call the method directly)
                destinationWorkbook.Combine(sourceWorkbook);
            }

            // Save the merged workbook – uses workbook-save rule
            destinationWorkbook.Save(outputFile, SaveFormat.Xlsx);
        }

        // Example usage
        public static void Main()
        {
            string[] filesToMerge = { "File1.xlsx", "File2.xlsx", "File3.xlsx" };
            string mergedOutput = "MergedWorkbook.xlsx";

            try
            {
                MergeWorkbooks(filesToMerge, mergedOutput);
                Console.WriteLine($"Workbooks merged successfully into '{mergedOutput}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during merging: {ex.Message}");
            }
        }
    }
}