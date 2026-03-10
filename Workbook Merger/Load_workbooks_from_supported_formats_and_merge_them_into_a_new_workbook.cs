using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsMergeExample
{
    class Program
    {
        static void MergeWorkbooks(string[] inputFiles, string outputFile)
        {
            // Destination workbook (empty)
            Workbook destWorkbook = new Workbook();

            foreach (string filePath in inputFiles)
            {
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"Warning: File not found – skipping '{filePath}'.");
                    continue;
                }

                // Load source workbook
                Workbook srcWorkbook = new Workbook(filePath);

                // Combine source into destination
                destWorkbook.Combine(srcWorkbook);
            }

            // Save merged workbook
            destWorkbook.Save(outputFile, SaveFormat.Xlsx);
        }

        static void Main(string[] args)
        {
            // Example input files (replace with actual paths as needed)
            string[] filesToMerge = new string[]
            {
                "File1.xlsx",
                "File2.xlsx",
                "File3.xlsx"
            };

            // Output file path
            string mergedOutput = "MergedWorkbook.xlsx";

            // Perform the merge operation
            MergeWorkbooks(filesToMerge, mergedOutput);

            Console.WriteLine($"Workbooks merged successfully into '{mergedOutput}'.");
        }
    }
}