using Novacode;
using System.Diagnostics;
using System.Collections.Generic;

namespace KIID.it.valuelab.hedgeinvest.KIID
{
    class CreateKIID
    {
        static void Main()
        {
            string fileName = @"D:\LAVORO\PROGETTI\HEDGEINVEST\KKID\DOC\TEMPLATE\TEMPLATE.docx";
            string outName  = @"D:\LAVORO\PROGETTI\HEDGEINVEST\KKID\DOC\OUT\OUT.docx";

            // Load a document in memory:
            DocX doc = DocX.Load(fileName);

            #region Grafico
            BarChart c = new BarChart();
            
            Series s1 = new Series("Andamento fondo");
            s1.Color = System.Drawing.Color.AntiqueWhite;
            List<string> years = new List<string>();
            years.Add("2013");
            years.Add("2014");
            years.Add("2015");
            years.Add("2016");
            List<string> performances = new List<string>();
            performances.Add("1.5");
            performances.Add("1.2");
            performances.Add("-4.55");
            performances.Add("-3");
            s1.Bind(years, performances);
            c.AddLegend(ChartLegendPosition.Bottom, false);
            c.AddSeries(s1);
            doc.InsertChart(c);
            #endregion

            //doc.ReplaceText("@ISIN@", "AAAAA");

            // Save to the output directory:
            doc.SaveAs(outName);

            // Open in Word:
            Process.Start("WINWORD.EXE", outName);
        }
    }
}
