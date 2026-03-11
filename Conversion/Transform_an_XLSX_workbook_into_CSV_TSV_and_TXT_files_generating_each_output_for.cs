using System;
using Aspose.Cells;

namespace AsposeCellsConversionDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source XLSX workbook
            string sourcePath = "source.xlsx";

            // Load the workbook (lifecycle rule: workbook-load)
            Workbook workbook = new Workbook(sourcePath);

            // Save as CSV (comma‑separated values)
            string csvPath = "output.csv";
            workbook.Save(csvPath, SaveFormat.Csv);   // lifecycle rule: workbook-save

            // Save as TSV (tab‑separated values)
            string tsvPath = "output.tsv";
            workbook.Save(tsvPath, SaveFormat.Tsv);   // lifecycle rule: workbook-save

            // Save as TXT – using CSV format but with a .txt extension
            string txtPath = "output.txt";
            workbook.Save(txtPath, SaveFormat.Csv);   // lifecycle rule: workbook-save

            Console.WriteLine("Conversion completed:");
            Console.WriteLine($"CSV  -> {csvPath}");
            Console.WriteLine($"TSV  -> {tsvPath}");
            Console.WriteLine($"TXT  -> {txtPath}");
        }
    }
}