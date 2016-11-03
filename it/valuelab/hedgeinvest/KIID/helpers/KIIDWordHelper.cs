using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using it.valuelab.hedgeinvest.helpers;
using System;

namespace KIID.it.valuelab.hedgeinvest.KIID.helpers
{
    class KIIDWordHelper : WordHelper
    {

        public KIIDWordHelper(String filename, String outName) : base(filename, outName) { }

        public void InsertProfiloRischio(string profiloRischio)
        {
            foreach (Table t in Document.MainDocumentPart.Document.Body.Elements<Table>())
            {
                TableRow row = t.Elements<TableRow>().ElementAt(4);
                foreach (TableCell cell in row.Elements<TableCell>())
                {
                    foreach (Table innerTable in cell.Elements<Table>())
                    {
                        TableRow innerRow = innerTable.Elements<TableRow>().ElementAt(0);
                        foreach (TableCell innerCell in innerRow.Elements<TableCell>())
                        {
                            if (innerCell.InnerText.Equals(profiloRischio))
                            {
                                innerCell.TableCellProperties.Shading.Fill= "CC9900";
                            }
                        }

                    }

                }    
            }
        }
    }


}
