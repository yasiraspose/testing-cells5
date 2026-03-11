using System;
using Aspose.Cells;
using Aspose.Cells.Slides;

class ExcelToPptxConverter
{
    static void Main()
    {
        // Path to the source Excel file
        string sourcePath = "input.xlsx";

        // Path where the resulting PowerPoint file will be saved
        string destinationPath = "output.pptx";

        // Load the Excel workbook (uses workbook-load rule)
        Workbook workbook = new Workbook(sourcePath);

        // Create PPTX save options (default constructor)
        PptxSaveOptions saveOptions = new PptxSaveOptions();

        // Optional: set the view type for the exported slides
        saveOptions.ExportViewType = SlideViewType.View;

        // Save the workbook as a PPTX presentation (uses workbook-save rule)
        workbook.Save(destinationPath, saveOptions);

        Console.WriteLine("Excel workbook has been successfully converted to PPTX.");
    }
}