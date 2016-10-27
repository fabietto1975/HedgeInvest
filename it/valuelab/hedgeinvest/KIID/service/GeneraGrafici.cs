using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Controls;

namespace AggiornaPowerpoint.c_classi
{
    internal static class GeneraGrafici
    {
        internal static string GeneraTorta(string nomeSlide, Dictionary<string, double> values, int nChart)
        {

            values.Remove("Totale"); //tolgo la l'ultima riga (il totale)
            ModificaGrafico modificaTorta = new ModificaTorta();
            modificaTorta.valori = values;
            string esito = modificaTorta.modifica_Grafico(App.pathDocPP, nomeSlide, "-", nChart, 0);
            return esito;
        }
        internal static string GeneraBarre(string nomeSlide, Dictionary<string, double> values, int nChart)
        {
            values.Remove("Totale"); //tolgo la l'ultima riga (il totale)
            ModificaGrafico modificaBarre = new ModificaBarre();
            modificaBarre.valori = values;
            string esito = modificaBarre.modifica_Grafico(App.pathDocPP, nomeSlide, "-", nChart, 0);
            return esito;
        }
        internal static string GeneraLe2Serie(string nomeSlide, List<Nav4colonne> values, string nomeFondo, string nomeBench)
        {
            ModificaGrafico modificaLinea = new ModificaLineaUcits();

            DictDaNav4Col dictValues = new serie1(values);
            modificaLinea.valori = dictValues;
            string esito = modificaLinea.modifica_Grafico(App.pathDocPP, nomeSlide, nomeFondo, 0, 1);
            esito += Environment.NewLine;

            dictValues = new serie3(values);
            modificaLinea.valori = dictValues;
            esito += modificaLinea.modifica_Grafico(App.pathDocPP, nomeSlide, nomeBench, 0, 2);
            esito += Environment.NewLine;

            return esito;
        }
        internal static string GeneraLe4Serie(string nomeSlide, List<Nav4colonne> values)
        {
            ModificaGrafico modificaLinea = new ModificaLinea();

            DictDaNav4Col dictValues = new serie1(values);
            modificaLinea.valori = dictValues;
            string esito = modificaLinea.modifica_Grafico(App.pathDocPP, nomeSlide, "HIGF", 0, 1);
            esito += Environment.NewLine;

            dictValues = new serie2(values);
            modificaLinea.valori = dictValues;
            esito += modificaLinea.modifica_Grafico(App.pathDocPP, nomeSlide, "HISS", 0, 2);
            esito += Environment.NewLine;

            dictValues = new serie3(values);
            modificaLinea.valori = dictValues;
            esito += modificaLinea.modifica_Grafico(App.pathDocPP, nomeSlide, "MSCI World in LC", 0, 3);
            esito += Environment.NewLine;

            dictValues = new serie4(values);
            modificaLinea.valori = dictValues;
            esito += modificaLinea.modifica_Grafico(App.pathDocPP, nomeSlide, "JP Morgan GBI in LC", 0, 4);
            esito += Environment.NewLine;
            return esito;
        }
        internal static string AggiornaGraficoLinea(string nomeSlide, DataGrid datagridRendimenti, string nomeFondo, string nomeBench)
        {
            string esito = string.Empty;
            if (datagridRendimenti.ItemsSource != null)
            {
                List<Nav4colonne> myDict = (List<Nav4colonne>)datagridRendimenti.ItemsSource;
                esito = GeneraGrafici.GeneraLe2Serie(nomeSlide, myDict, nomeFondo, nomeBench);
            }
            else
            {
                esito = string.Format("Generare i dati per '{0}' !", nomeSlide);
            }
            return esito + Environment.NewLine;
        }

    }
}
