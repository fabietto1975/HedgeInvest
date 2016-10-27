using DocumentFormat.OpenXml.Drawing.Charts;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AggiornaPowerpoint.c_classi
{
    public class GraficoBarre
    {
        public ChartSpace BarChartSpace { get; set; }
        public Object Barre { get; set; }
        public BarChartSeries barChartSeries1;
        public virtual void getBarre() { }
    }

    class GraficoBarre2d : GraficoBarre
    {      
        public override void getBarre()
        {
            BarChart barChart = BarChartSpace.Descendants<BarChart>().FirstOrDefault();
            barChartSeries1 = barChart.Descendants<BarChartSeries>().FirstOrDefault();
            Barre = barChart;
        }
    }
}

