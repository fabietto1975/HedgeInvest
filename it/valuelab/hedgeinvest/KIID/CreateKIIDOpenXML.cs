using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using s = KIID.it.valuelab.hedgeinvest.KIID.service;
using System.Diagnostics;
using System.IO;
using d = DocumentFormat.OpenXml.Drawing;
using dc = DocumentFormat.OpenXml.Drawing.Charts;
using dw = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace KIID.it.valuelab.hedgeinvest.KIID
{
    class CreateKIIDOpenXML
    {
        static void Main()
        {
            s.KIIDService service = new s.KIIDService();
            service.readFundsData();

            /*
            string template = @"D:\LAVORO\PROGETTI\HEDGEINVEST\KKID\DOC\TEMPLATE\TEMPLATE.docx";
            string outName = @"D:\LAVORO\PROGETTI\HEDGEINVEST\KKID\DOC\OUT\OUT.docx";

            byte[] byteArray = File.ReadAllBytes(template);
            using (MemoryStream stream = new MemoryStream())
            {
                stream.Write(byteArray, 0, (int)byteArray.Length);
                using (WordprocessingDocument wordDocument =
                    WordprocessingDocument.Open(stream, true)) 
                {
                    Body body = wordDocument.MainDocumentPart.Document.Body;

                    ChartPart chartPart = wordDocument.MainDocumentPart.AddNewPart<ChartPart>("fundPerformance");

                    dc.BarChart barChart = new dc.BarChart();

                    Drawing dr = new Drawing();
                    dw.Inline inline = new dw.Inline();
                    dw.Extent extent = new dw.Extent();
                    extent.Cx = 10000;
                    extent.Cy = 2342342;
                    inline.Extent = extent;
                    dr.Inline = inline;

                    d.Graphic g = new d.Graphic();
                    d.GraphicData gd = new d.GraphicData();
                    inline.Graphic = g;


                }

                File.WriteAllBytes(outName, stream.ToArray());
            }

            // Open in Word:
            Process.Start("WINWORD.EXE", outName);
            */

        }
    }
}
