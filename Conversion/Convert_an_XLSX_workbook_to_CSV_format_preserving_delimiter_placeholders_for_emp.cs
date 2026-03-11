using System;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsConversion
{
    public class XlsxToCsvConverter
    {
        public static void Convert(string inputPath, string outputPath)
        {
            Workbook workbook = new Workbook(inputPath);
            TxtSaveOptions csvOptions = new TxtSaveOptions
            {
                Encoding = Encoding.UTF8,
                KeepSeparatorsForBlankRow = true
            };
            workbook.Save(outputPath, csvOptions);
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: AsposeCellsRunner <input.xlsx> <output.csv>");
                return;
            }

            XlsxToCsvConverter.Convert(args[0], args[1]);
            Console.WriteLine("Conversion completed.");
        }
    }
}