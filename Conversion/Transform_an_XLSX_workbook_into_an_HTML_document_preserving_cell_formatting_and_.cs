using System;
using Aspose.Cells;

class ExcelToHtmlConverter
{
    static void Main()
    {
        // Paths for source Excel file and destination HTML file
        string inputPath = "input.xlsx";
        string outputPath = "output.html";

        // Load the workbook from the XLSX file
        Workbook workbook = new Workbook(inputPath); // workbook-load

        // Set up HTML save options to keep formatting and properties
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.ExportWorkbookProperties = true;          // keep workbook metadata
        htmlOptions.ExportWorksheetProperties = true;        // keep worksheet metadata
        htmlOptions.ExportGridLines = true;                  // include grid lines in HTML
        htmlOptions.HtmlVersion = HtmlVersion.Html5;         // generate modern HTML5

        // Save the workbook as an HTML document with the specified options
        workbook.Save(outputPath, htmlOptions); // workbook-save
    }
}