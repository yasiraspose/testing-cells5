using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class AddSparklineToMacroTemplate
{
    static void Main()
    {
        // Paths for the macro‑enabled template and the output file
        string templatePath = "Template.xltm";
        string outputPath   = "Result.xltm";

        // Load the existing XLTM template (workbook-load rule)
        Workbook workbook = new Workbook(templatePath);

        // Access the first worksheet (worksheet-access rule)
        Worksheet sheet = workbook.Worksheets[0];

        // Populate some sample data that the sparkline will represent
        sheet.Cells["A1"].PutValue(5);
        sheet.Cells["B1"].PutValue(2);
        sheet.Cells["C1"].PutValue(1);
        sheet.Cells["D1"].PutValue(3);

        // Define the cell where the sparkline will be placed (E1)
        CellArea sparklineLocation = new CellArea
        {
            StartRow    = 0, // Row 1 (zero‑based)
            EndRow      = 0,
            StartColumn = 4, // Column E (zero‑based)
            EndColumn   = 4
        };

        // Add a sparkline group to the worksheet (no specific rule, free‑form code)
        int groupIndex = sheet.SparklineGroups.Add(
            SparklineType.Line,                 // Type of sparkline
            sheet.Name + "!A1:D1",              // Data range for the sparkline
            false,                              // Plot by row (horizontal)
            sparklineLocation);                 // Where the sparkline appears

        SparklineGroup group = sheet.SparklineGroups[groupIndex];

        // Optional: apply a preset style to the sparkline group
        group.PresetStyle = SparklinePresetStyleType.Style5;

        // Save the workbook as a macro‑enabled template (workbook-save rule)
        workbook.Save(outputPath, SaveFormat.Xltm);
    }
}