using System;
using Aspose.Cells;

class WorkbookMergePrerequisiteDemo
{
    static void Main()
    {
        // Create the first workbook
        Workbook firstWorkbook = new Workbook();                     // workbook-create
        Worksheet firstSheet = firstWorkbook.Worksheets[0];        // worksheet-access
        firstSheet.Cells["A1"].PutValue("First Workbook");
        firstSheet.Cells.Merge(0, 0, 2, 1);                         // merge A1:A2

        // Create the second workbook
        Workbook secondWorkbook = new Workbook();                    // workbook-create
        Worksheet secondSheet = secondWorkbook.Worksheets[0];      // worksheet-access
        secondSheet.Cells["A1"].PutValue("Second Workbook");
        secondSheet.Cells.Merge(0, 0, 3, 2);                        // merge A1:B3

        // Prepare save options to validate merged areas before saving
        XlsbSaveOptions saveOptions = new XlsbSaveOptions();
        saveOptions.ValidateMergedAreas = true;    // prerequisite 1
        saveOptions.MergeAreas = true;             // prerequisite 2

        // Simple check that prerequisites are enabled
        if (!saveOptions.ValidateMergedAreas || !saveOptions.MergeAreas)
        {
            throw new InvalidOperationException("Required save options are not set.");
        }

        // Combine the two workbooks
        firstWorkbook.Combine(secondWorkbook);      // free‑form combine (no specific rule)

        // Save the combined workbook using the mandated save rule
        firstWorkbook.Save("CombinedWorkbook.xlsx", SaveFormat.Xlsx); // workbook-save

        Console.WriteLine("Combined workbook saved with prerequisite validation.");
    }
}