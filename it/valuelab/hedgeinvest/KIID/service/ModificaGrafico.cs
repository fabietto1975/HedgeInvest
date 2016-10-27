using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Linq;

namespace AggiornaPowerpoint.c_classi
{
    class ModificaGrafico
    {
        public Dictionary<string, double> valori { get; set; }

        public string modifica_Grafico(string path, string nomeSlide, string titoloSerie, int nChart, int nSerie)
        {
            PresentationDocument presentationDoc = null;
            try
            {
                presentationDoc = PresentationDocument.Open(path, true);

                SlidePart theSlidePart = PPointDocument.getSlidePart(presentationDoc, nomeSlide);

                if (theSlidePart != null)
                {
                    if (theSlidePart.ChartParts.Count() > 0)
                    {
                        ChartPart chartPart = theSlidePart.ChartParts.Skip(nChart).FirstOrDefault();

                        if (ReplaceValuesInChartInSlide(chartPart, titoloSerie, nSerie) == "ok")
                        {
                            presentationDoc.PresentationPart.Presentation.Save();
                            return string.Format("ok grafico {0} {1}", nomeSlide, titoloSerie);
                        }
                        else
                        {
                            return string.Format("non trovo il grafico nella slide '{0}'", nomeSlide);
                        }

                    }
                    else
                    {
                        return string.Format("non trovo nessun grafico nella slide '{0}'", nomeSlide);
                    }

                }
                else
                {
                    return string.Format("non trovo la slide '{0}'", nomeSlide);
                }
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
            finally
            {
                if (presentationDoc != null)
                {
                    presentationDoc.Dispose();
                }
            }
        }

        protected void modificaChartData(string FormatoValori, string titoloSerie, out SeriesText seriesText1, out CategoryAxisData categoryAxisData1, out Values values1)
        {
            seriesText1 = new SeriesText();

            StringReference stringReference1 = new StringReference();
            Formula formula1 = new Formula();
            formula1.Text = "Foglio1!$B$1";

            StringCache stringCache1 = new StringCache();
            PointCount pointCount1 = new PointCount() { Val = (UInt32Value)1U };

            StringPoint stringPoint1 = new StringPoint() { Index = (UInt32Value)0U };
            NumericValue numericValue1 = new NumericValue();
            numericValue1.Text = titoloSerie;

            stringPoint1.Append(numericValue1);

            stringCache1.Append(pointCount1);
            stringCache1.Append(stringPoint1);

            stringReference1.Append(formula1);
            stringReference1.Append(stringCache1);

            seriesText1.Append(stringReference1);

            DataPoint dataPoint1 = new DataPoint();
            Index index2 = new Index() { Val = (UInt32Value)2U };


            dataPoint1.Append(index2);

            //################################i testi ####################################
            categoryAxisData1 = new CategoryAxisData();

            StringReference stringReference2 = new StringReference();
            Formula formula2 = new Formula();
            formula2.Text = string.Format("Foglio1!$A$2:$A${0}", valori.Count + 2);

            StringCache stringCache2 = new StringCache();
            UInt32Value nValori = Convert.ToUInt32(valori.Count);

            PointCount pointCount2 = new PointCount() { Val = nValori };

            StringPoint[] stringPoints = new StringPoint[nValori];
            UInt32Value n = 0;

            foreach (KeyValuePair<string, double> item in valori)
            {
                stringPoints[n] = new StringPoint() { Index = (UInt32Value)n };
                NumericValue numericValue2 = new NumericValue();
                numericValue2.Text = item.Key;
                stringPoints[n].Append(numericValue2);
                n += 1;
            }



            stringCache2.Append(pointCount2);
            for (int i = 0; i < n; i++)
            {
                stringCache2.Append(stringPoints[i]);
            }

            stringReference2.Append(formula2);
            stringReference2.Append(stringCache2);

            categoryAxisData1.Append(stringReference2);


            //################################i valori####################################

            values1 = new Values();

            NumberReference numberReference1 = new NumberReference();
            Formula formula3 = new Formula();
            formula3.Text = string.Format("Foglio1!$B$2:$B${0}", valori.Count + 2);

            NumberingCache numberingCache1 = new NumberingCache();
            FormatCode formatCode1 = new FormatCode();
            formatCode1.Text = FormatoValori; //<-----------------------------------------------------------
            PointCount pointCount3 = new PointCount() { Val = nValori };

            NumericPoint[] numericPoints = new NumericPoint[nValori];
            n = 0;
            foreach (KeyValuePair<string, double> item in valori)
            {
                numericPoints[n] = new NumericPoint() { Index = (UInt32Value)n };
                NumericValue numericValue = new NumericValue();
                numericValue.Text = item.Value.ToString();
                numericValue.Text = numericValue.Text.Replace(",", "."); // devo forzare il formato americano
                numericPoints[n].Append(numericValue);
                n += 1;
            }



            numberingCache1.Append(formatCode1);
            numberingCache1.Append(pointCount3);
            for (int i = 0; i < n; i++)
            {
                numberingCache1.Append(numericPoints[i]);
            }



            numberReference1.Append(formula3);
            numberReference1.Append(numberingCache1);

            values1.Append(numberReference1);
        }

        protected virtual string ReplaceValuesInChartInSlide(ChartPart chartPart, string categoryTitle, int nSerie)
        {
            return string.Empty;
        }

    }
}
