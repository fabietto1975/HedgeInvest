using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;

namespace AggiornaPowerpoint.c_classi
{
    class ModificaBarre : ModificaGrafico
    {
        protected override string ReplaceValuesInChartInSlide(ChartPart chartPart,  string categoryTitle,int nSerie)
        {
            ChartSpace chartSpace = chartPart.ChartSpace;

            GraficoBarre myBarre = new GraficoBarre2d();
                      
            myBarre.BarChartSpace = chartSpace;
            myBarre.getBarre();

            if (myBarre.Barre != null)
            {             

                BarChartSeries barChartSeries1 = myBarre.barChartSeries1;

                SeriesText seriesText1;
                CategoryAxisData categoryAxisData1;
                Values values1;

                modificaChartData("0,0","", out seriesText1, out categoryAxisData1, out values1);

                barChartSeries1.SeriesText = seriesText1;
                barChartSeries1.Append(categoryAxisData1);
                barChartSeries1.Append(values1);
               
                return "ok";
            }
            else
            {
                return "non trovo il grafico a barre!";
            }
        }      
    }
}
