using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using it.valuelab.hedgeinvest.helpers;
using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Drawing.Charts;
using System.Globalization;

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
                                innerCell.TableCellProperties.Shading.Fill= "CC9900";
                            }
                        }
                    }
                }    
            }
        }

        private void fillPoints(string baseFormula, String mode, List<String> data)
        {
            int idx = data.Count;

            ChartPart cp = Document.MainDocumentPart.ChartParts.FirstOrDefault();
            Chart chart = cp.ChartSpace.Elements<Chart>().FirstOrDefault();
            BarChart barchart = chart.PlotArea.Elements<BarChart>().FirstOrDefault();
            BarChartSeries series = barchart.Elements<BarChartSeries>().FirstOrDefault();

            CategoryAxisData labels = new CategoryAxisData();
            Values values = new Values();

            NumberReference nref = new NumberReference();
            Formula f = new Formula(baseFormula + idx);
            nref.Formula = f;
            NumberingCache nc = new NumberingCache();//nref.Descendants<NumberingCache>().First();
            nc.PointCount = new PointCount();
            nc.PointCount.Val = (uint)idx;
            int pointIndex = 1;
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
                series.AppendChild(labels);
            }
            else if ("VALUES".Equals(mode))
            {
                values.AppendChild(nref);
                series.AppendChild(values);
            };



        }

        public void EditPerformanceTable(SortedDictionary<string, string> performances)
        {
            if (performances != null)
            {
                fillPoints("Foglio1!$A$2:$A$", "LABELS", performances.Keys.ToList());
                fillPoints("Foglio1!$B$2:$B$", "VALUES", performances.Values.ToList());
            }
            else
            {
                
            }
    
        }
    }
}
