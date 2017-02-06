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
            Log.Info(formula);
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
                    if (val != "0")
                    {
                        float valuePerc = float.Parse(val, CultureInfo.InvariantCulture.NumberFormat) * 100;
                        value.Text = valuePerc.ToString(CultureInfo.InvariantCulture);
                    }
                    else
                    {
                        value.Text = "";

                    }
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

                        int cnt = 0;
                        //Cosi facendo ci si limita al numero di righe predisposte dal template, max 5
                        int rowCount = performances.Keys.Count();
                        
                        foreach (Row row in sd.Elements<Row>())
                        {
                            Log.Debug("cnt: " + cnt);
                            if (cnt > 0)
                            {
                                string label = "";
                                string value = "";
                                if (cnt <= rowCount)
                                {
                                    label = performances.Keys.ElementAt((cnt - 1));
                                    value = (float.Parse(performances[label], CultureInfo.InvariantCulture.NumberFormat) * 100).ToString(CultureInfo.InvariantCulture); 
                                    if (value.Equals("0") || value.Equals("") || value == null)
                                    {
                                        value = "";
                                    }

                                }
                                Log.Debug(label + " : " + (value==null));
                                if (row.Elements<Cell>() != null && row.Elements<Cell>().Count() > 0)
                                {
                                    row.Elements<Cell>().ElementAt(0).Elements<CellValue>().FirstOrDefault().Text = label;
                                    if (value != null)
                                    {
                                        Log.Debug("Imposto valore " + value);
                                        row.Elements<Cell>().ElementAt(1).Elements<CellValue>().FirstOrDefault().Text = value;
                                    }
                                    Log.Debug("Cella esistente");
                                }
                                else
                                {
                                    Cell labelCell = new Cell();
                                    labelCell.CellValue = new CellValue();
                                    labelCell.CellValue.Text = label;
                                    Cell valueCell = new Cell();
                                    valueCell.CellValue = new CellValue();
                                    if (value != null)
                                    {
                                        Log.Debug("Imposto valore " + value);
                                        valueCell.CellValue.Text = value;
                                    }
                                    row.AppendChild<Cell>(labelCell);
                                    row.AppendChild<Cell>(valueCell);
                                    Log.Debug("Cella presente");
                                }

                            }
                            cnt++;
                        }
                        Log.Debug(sd.InnerXml);
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
