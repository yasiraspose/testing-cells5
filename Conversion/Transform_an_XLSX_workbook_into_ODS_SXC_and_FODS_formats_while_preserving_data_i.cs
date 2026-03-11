using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Paths for source and target files
        string sourcePath = "input.xlsx";
        string odsPath = "output.ods";
        string sxcPath = "output.sxc";
        string fodsPath = "output.fods";

        // Load the existing XLSX workbook
        Workbook workbook = new Workbook(sourcePath); // workbook-load

        // Save the workbook in OpenDocument Spreadsheet format (ODS)
        workbook.Save(odsPath, SaveFormat.Ods); // workbook-save

        // Save the workbook in StarOffice Calc format (SXC)
        workbook.Save(sxcPath, SaveFormat.Sxc); // workbook-save

        // Save the workbook in OpenDocument Flat XML Spreadsheet format (FODS)
        workbook.Save(fodsPath, SaveFormat.Fods); // workbook-save
    }
}