using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Sample SpreadsheetML that defines a Timeline control.
        string timelineXml = @"<?xml version='1.0' encoding='UTF-8'?>
<worksheet xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
           xmlns:xdr='http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'
           xmlns:a='http://schemas.openxmlformats.org/drawingml/2006/main'>
  <drawing>
    <xdr:wsDr>
      <xdr:twoCellAnchor>
        <xdr:from>
          <xdr:col>0</xdr:col>
          <xdr:colOff>0</xdr:colOff>
          <xdr:row>0</xdr:row>
          <xdr:rowOff>0</xdr:rowOff>
        </xdr:from>
        <xdr:to>
          <xdr:col>5</xdr:col>
          <xdr:colOff>0</xdr:colOff>
          <xdr:row>10</xdr:row>
          <xdr:rowOff>0</xdr:rowOff>
        </xdr:to>
        <xdr:graphicFrame macro=''>
          <xdr:nvGraphicFramePr>
            <xdr:cNvPr id='2' name='Timeline 1'/>
            <xdr:cNvGraphicFramePr/>
          </xdr:nvGraphicFramePr>
          <xdr:xfrm>
            <a:off x='0' y='0'/>
            <a:ext cx='0' cy='0'/>
          </xdr:xfrm>
          <a:graphic>
            <a:graphicData uri='http://schemas.microsoft.com/office/spreadsheetml/2009/9/main'>
              <timeline xmlns='http://schemas.microsoft.com/office/spreadsheetml/2009/9/main'
                        name='MyTimeline'
                        baseFieldName='Date'
                        pivotTableId='0'/>
            </a:graphicData>
          </a:graphic>
        </xdr:graphicFrame>
        <xdr:clientData/>
      </xdr:twoCellAnchor>
    </xdr:wsDr>
  </drawing>
</worksheet>";

        using (MemoryStream xmlStream = new MemoryStream())
        {
            using (StreamWriter writer = new StreamWriter(xmlStream))
            {
                writer.Write(timelineXml);
                writer.Flush();
                xmlStream.Position = 0;

                workbook.ImportXml(xmlStream, "Sheet1", 0, 0);
            }
        }

        workbook.Save("TimelineWorkbook.xlsx", SaveFormat.Xlsx);
    }
}