using Aspose.Cells;
using System.Net.Mime;
using Microsoft.AspNetCore.Http;
using System.IO;

public class WorkbookExport
{
    public void ExportToHttpResponse(HttpResponse response)
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add sample data
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Hello");
        sheet.Cells["B1"].PutValue("World");

        // Define the file name for the download
        string fileName = "Report.xlsx";

        // Set response headers
        response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
        ContentDisposition disposition = new ContentDisposition
        {
            FileName = fileName,
            Inline = false
        };
        response.Headers["Content-Disposition"] = disposition.ToString();

        // Save workbook to response stream
        using (MemoryStream ms = new MemoryStream())
        {
            workbook.Save(ms, SaveFormat.Xlsx);
            ms.Position = 0;
            ms.CopyTo(response.Body);
        }
    }
}