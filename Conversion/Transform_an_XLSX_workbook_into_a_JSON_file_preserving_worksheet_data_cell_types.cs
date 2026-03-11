using System;
using Aspose.Cells;

public class ExcelToJsonConverter
{
    public static void Convert(string inputPath, string outputPath)
    {
        Workbook workbook = new Workbook(inputPath);
        JsonSaveOptions jsonOptions = new JsonSaveOptions
        {
            AlwaysExportAsJsonObject = true,
            ToExcelStruct = true,
            ExportEmptyCells = true,
            ExportNestedStructure = true,
            ExportAsString = true,
            SkipEmptyRows = false
        };
        workbook.Save(outputPath, jsonOptions);
    }

    public static void Main(string[] args)
    {
        if (args.Length < 2)
        {
            Console.WriteLine("Usage: ExcelToJsonConverter <input.xlsx> <output.json>");
            return;
        }

        Convert(args[0], args[1]);
        Console.WriteLine("Conversion completed.");
    }
}