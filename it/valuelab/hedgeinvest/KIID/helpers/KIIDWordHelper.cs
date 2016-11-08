using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using it.valuelab.hedgeinvest.helpers;
using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using System.IO;
using System.Text;
using DocumentFormat.OpenXml;

namespace KIID.it.valuelab.hedgeinvest.KIID.helpers
{
    class KIIDWordHelper : WordHelper
    {

        public KIIDWordHelper(String filename, String outName) : base(filename, outName) { }

        public void InsertProfiloRischio(string profiloRischio)
        {
            foreach (Table t in Document.MainDocumentPart.Document.Body.Elements<Table>())
            {
                TableRow row = t.Elements<TableRow>().ElementAt(4); //Sezione "Profilo di rischio e di rendimento"
                foreach (TableCell cell in row.Elements<TableCell>())
                {
                    foreach (Table innerTable in cell.Elements<Table>())
                    {
                        TableRow innerRow = innerTable.Elements<TableRow>().ElementAt(0);
                        foreach (TableCell innerCell in innerRow.Elements<TableCell>())
                        {
                            if (innerCell.InnerText.Equals(profiloRischio))
                            {
                                innerCell.TableCellProperties.Shading.Fill = "CC9900";
                            }
                        }
                    }
                }
            }
        }

        public void InsertPerformanceTable(SortedDictionary<string, string> performances)
        {
            foreach (Table t in Document.MainDocumentPart.Document.Body.Elements<Table>())
            {
                TableRow row = t.Elements<TableRow>().ElementAt(9);//Sezione "Risultati ottenuti nel passato"
                TableCell cell = row.Elements<TableCell>().ElementAt(0);

                Drawing d = cell.Elements<Paragraph>().ElementAt(0).Elements<Run>().ElementAt(0).Elements<Drawing>().ElementAt(0);
                System.Diagnostics.Debug.WriteLine(d.InnerText);
                Inline inline = d.Inline;
            }
        }

        public void replaceText2(string oldtext, string newtext)
        {


            string altChunkId = oldtext.Replace("@", "");
            MainDocumentPart mainDocPart = Document.MainDocumentPart;


            AlternativeFormatImportPart chunk = mainDocPart.AddAlternativeFormatImportPart(AlternativeFormatImportPartType.Rtf, altChunkId);
            string rtfEncodedString = newtext;

            using (MemoryStream ms = new MemoryStream(Encoding.ASCII.GetBytes(rtfEncodedString)))
            {
                chunk.FeedData(ms);
            }

            AltChunk altChunk = new AltChunk();
            altChunk.Id = altChunkId;

            //mainDocPart.Document.Body.InsertAfter(
            //  altChunk, mainDocPart.Document.Body.Elements<Paragraph>().Last());

            //mainDocPart.Document.Save();

            //Document.MainDocumentPart.Document.Descendants<SdtRun>()
            var blocks = from t in Document.MainDocumentPart.Document.Body.Elements<Table>()
                         //from s in t.Elements<Run>()
                         //where (
                         //        s.GetFirstChild<SdtProperties>().GetFirstChild<Text>() != null
                         //        && s.GetFirstChild<SdtProperties>().GetFirstChild<Text>().ToString().Contains(oldtext)
                         //        )
                         select t;

            var sdt = blocks.FirstOrDefault();

            OpenXmlElement parent = sdt.Parent;

            parent.InsertAfter(altChunk, blocks.SingleOrDefault());

            sdt.Remove();

            Document.MainDocumentPart.Document.Save();
        }
    }
}
