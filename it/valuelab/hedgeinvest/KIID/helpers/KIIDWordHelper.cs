using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using it.valuelab.hedgeinvest.helpers;
using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Drawing.Charts;
using System.Globalization;
using System.IO;
using DocumentFormat.OpenXml.Spreadsheet;
using T = DocumentFormat.OpenXml.Wordprocessing.Table;

namespace KIID.it.valuelab.hedgeinvest.KIID.helpers
{
    class KIIDWordHelper : WordHelper
    {

        public KIIDWordHelper(string filename, String outName) : base(filename, outName) { }

        public void InsertProfiloRischio(string profiloRischio)
        {

            T table = FindByCaption("TABELLACLASSEDIRISCHIO");

            TableRow innerRow = table.Elements<TableRow>().ElementAt(0);

            foreach (TableCell innerCell in innerRow.Elements<TableCell>())
            {
                if (innerCell.InnerText==profiloRischio)
                {
                    innerCell.TableCellProperties.Shading.Fill = "CC9900";
                }
            }

        }


        private void CopyStream(Stream input, Stream output)
        {
            byte[] buffer = new byte[32768];
            while (true)
            {
                int read = input.Read(buffer, 0, buffer.Length);
                if (read <= 0)
                    return;
                output.Write(buffer, 0, read);
            }
        }

        private void FillPoints(string baseFormula, String mode, List<String> data)
        {
            int idx = data.Count;

            ChartPart cp = Document.MainDocumentPart.ChartParts.FirstOrDefault();
            Chart chart = cp.ChartSpace.Elements<Chart>().FirstOrDefault();
            BarChart barchart = chart.PlotArea.Elements<BarChart>().FirstOrDefault();

            BarChartSeries series = barchart.Elements<BarChartSeries>().FirstOrDefault();
            
            CategoryAxisData labels = new CategoryAxisData();
            DocumentFormat.OpenXml.Drawing.Charts.Values values = new DocumentFormat.OpenXml.Drawing.Charts.Values();

            NumberReference nref = new NumberReference();
            string formula = baseFormula + (idx + 1);
            DocumentFormat.OpenXml.Drawing.Charts.Formula f = new DocumentFormat.OpenXml.Drawing.Charts.Formula();
            f.Text = formula;
            nref.Formula = f;
            NumberingCache nc = new NumberingCache();//nref.Descendants<NumberingCache>().First();
            nc.PointCount = new PointCount();
            nc.PointCount.Val = (uint)idx;
            int pointIndex = 0;
            foreach (string val in data)
            {
                NumericPoint point = new NumericPoint();
                point.Index = (uint)pointIndex;
                NumericValue value = new NumericValue();
                if ("LABELS".Equals(mode))
                {
                    value.Text = val;
                }
                else if ("VALUES".Equals(mode))
                {
                    float valuePerc = float.Parse(val, CultureInfo.InvariantCulture.NumberFormat) * 100;
                    value.Text = valuePerc.ToString(CultureInfo.InvariantCulture);
                }
                point.AppendChild(value);
                nc.AppendChild(point);
                pointIndex++;
            }
            nref.AppendChild(nc);

            if ("LABELS".Equals(mode))
            {
                labels.AppendChild(nref);
                series.ReplaceChild<CategoryAxisData>(labels, series.Elements<CategoryAxisData>().FirstOrDefault());
                
            }
            else if ("VALUES".Equals(mode))
            {
                values.AppendChild(nref);
                series.ReplaceChild<DocumentFormat.OpenXml.Drawing.Charts.Values>(values, series.Elements< DocumentFormat.OpenXml.Drawing.Charts.Values>().FirstOrDefault());
            };
        }

        public void EditPerformanceChart(SortedDictionary<string, string> performances)
        {
            if (performances != null)
            {
                /*
                foreach (DocumentFormat.OpenXml.Wordprocessing.Table t in Document.MainDocumentPart.Document.Body.Elements<DocumentFormat.OpenXml.Wordprocessing.Table>())
                {
                    TableRow row = t.Elements<TableRow>().ElementAt(10);
                    row.Remove();
                }
                */
                this.RemoveRowByContent("@TESTO2@");


                //Aggiornamento XML
                FillPoints("Foglio1!$A$2:$A$", "LABELS", performances.Keys.ToList());
                FillPoints("Foglio1!$B$2:$B$", "VALUES", performances.Values.ToList());

                //Aggiornamento Embedded XLS
                ChartPart cp = Document.MainDocumentPart.ChartParts.FirstOrDefault();
                ExternalData ed = cp.ChartSpace.Elements<ExternalData>().FirstOrDefault();
                EmbeddedPackagePart epp = (EmbeddedPackagePart)cp.Parts.Where(
                            pt => pt.RelationshipId == ed.Id)
                                                                .FirstOrDefault()
                                                                .OpenXmlPart;
                using (System.IO.Stream str = epp.GetStream())
                using (MemoryStream ms = new MemoryStream())
                {
                    CopyStream(str, ms);
                    using (SpreadsheetDocument spreadsheetDoc = SpreadsheetDocument.Open(ms, true))
                    {
                        Sheet ws = (Sheet)spreadsheetDoc.WorkbookPart.Workbook.Sheets.FirstOrDefault();
                        string sheetId = ws.Id;
                        WorksheetPart wsp = (WorksheetPart)spreadsheetDoc.WorkbookPart.Parts
                                    .Where(pt => pt.RelationshipId == sheetId)
                                    .FirstOrDefault()
                                    .OpenXmlPart;
                        SheetData sd = wsp.Worksheet.Elements<SheetData>().FirstOrDefault();
                        //System.Diagnostics.Debug.WriteLine(sd.InnerXml);
                        //System.Diagnostics.Debug.WriteLine(mainRow.InnerXml);
                        //sd.RemoveAllChildren<Row>();
                        //mainRow.RemoveAllChildren();
                        int rowCount = sd.Elements<Row>().Count();
                        //System.Diagnostics.Debug.WriteLine(rowCount);
                        for (int idx = 0; idx < performances.Count; idx++)
                        {
                            Row row;
                            Boolean append = false;
                            if (idx + 1 < rowCount)
                            {
                                row = sd.Elements<Row>().ElementAt((idx + 1));
                            }
                            else
                            {
                                row = new Row();
                                append = true;

                            }

                            //new Row();
                            string key = performances.Keys.ElementAt(idx);
                            //Labels
                            Cell labelCell = new Cell();
                            labelCell.CellValue = new CellValue();
                            labelCell.Elements<CellValue>().FirstOrDefault().Text = key;

                            //Values
                            Cell valueCell = new Cell();
                            valueCell.CellValue = new CellValue();
                            float valuePerc = float.Parse(performances[key], CultureInfo.InvariantCulture.NumberFormat) * 100;
                            valueCell.Elements<CellValue>().FirstOrDefault().Text = valuePerc.ToString(CultureInfo.InvariantCulture);
                            row.AppendChild(labelCell);
                            row.Append(valueCell);

                            if (append)
                                sd.AppendChild(row);
                        }


                    }
                    using (Stream s = epp.GetStream())
                        ms.WriteTo(s);
                }
            }
            else
            {
                this.RemoveRowByContent("@TESTO3@");
            }
        }
    }
}
