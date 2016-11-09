using m = it.valuelab.hedgeinvest.KIID.model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using it.valuelab.hedgeinvest.helpers;
using KIID.it.valuelab.hedgeinvest.KIID.helpers;
using System.Globalization;

namespace it.valuelab.hedgeinvest.KIID.service
{
    public class KIIDService
    {
        private string template;
        private string datafile;
        private string outputfolder;
        private string language;
        private DateTime datagenerazione;

        public KIIDService(string _template, string _datafile, string _outputfolder, string _language, DateTime _datagenerazione)
        {
            template = _template;
            datafile = _datafile;
            outputfolder = _outputfolder;
            datagenerazione = _datagenerazione;
            language = _language;
        }

        public List<m.KIIDData> readFundsData()

        {
            const string mainSheetname = "DATI KIID";
            const string performanceSheetname = "PERFORMANCE";
            List<m.KIIDData> result = new List<m.KIIDData>();
            using (ExcelHelper excelHelper = new ExcelHelper(datafile))
            {
                //Performance

                int row = 2;
                Dictionary<string, SortedDictionary<string,string>> isinPerformanceAnnoMap = new Dictionary<string, SortedDictionary<string, string>>() ;
                string isin = excelHelper.GetValue(performanceSheetname, "B", row.ToString());
                while (!string.IsNullOrEmpty(isin))
                {
                    string anno = excelHelper.GetValue(performanceSheetname, "C", row.ToString());
                    string dato = excelHelper.GetValue(performanceSheetname, "D", row.ToString());
                    SortedDictionary<string, string> isinPerformanceAnno; 
                    if (!isinPerformanceAnnoMap.TryGetValue(isin, out isinPerformanceAnno))
                    {
                        isinPerformanceAnno = new SortedDictionary<string, string>();
                    }
                    isinPerformanceAnno[anno] = dato;
                    isinPerformanceAnnoMap[isin] =  isinPerformanceAnno;
                    row++;
                    isin = excelHelper.GetValue(performanceSheetname, "B", row.ToString());
                }
                string suffix = "";
                if (!language.Equals("IT"))
                {
                    suffix += " - " + language;
                }
                Dictionary<string, string> fieldPosition = new Dictionary<String, string>();
                //Header row
                string classeCol = excelHelper.GetCellColumn(excelHelper.GetCellByContent(mainSheetname, "CLASSE", 1));
                string isinCol = excelHelper.GetCellColumn(excelHelper.GetCellByContent(mainSheetname, "ISIN", 1));
                string classeDiRischioCol = excelHelper.GetCellColumn(excelHelper.GetCellByContent(mainSheetname, "CLASSE DI RISCHIO", 1));
                string testo1Col = excelHelper.GetCellColumn(excelHelper.GetCellByContent(mainSheetname, "TESTO1" + suffix, 1));
                string testo2Col = excelHelper.GetCellColumn(excelHelper.GetCellByContent(mainSheetname, "TESTO2" + suffix, 1));
                string testo3Col = excelHelper.GetCellColumn(excelHelper.GetCellByContent(mainSheetname, "TESTO3" + suffix, 1));
                string spesesottoscrizioneCol = excelHelper.GetCellColumn(excelHelper.GetCellByContent(mainSheetname, "SPESE SOTTOSCRIZIONE", 1));
                string speserimborsoCol = excelHelper.GetCellColumn(excelHelper.GetCellByContent(mainSheetname, "SPESE DI RIMBORSO", 1));
                string spesecorrentiCol = excelHelper.GetCellColumn(excelHelper.GetCellByContent(mainSheetname, "SPESE CORRENTI", 1));
                string spesediconversioneCol = excelHelper.GetCellColumn(excelHelper.GetCellByContent(mainSheetname, "SPESE DI CONVERSIONE", 1));
                string commissioniRendimentoCol = excelHelper.GetCellColumn(excelHelper.GetCellByContent(mainSheetname, "COMMISSIONI LEGATE AL RENDIMENTO" + suffix, 1));
                string informazionipraticheCol = excelHelper.GetCellColumn(excelHelper.GetCellByContent(mainSheetname, "INFORMAZIONI PRATICHE" + suffix, 1));
                string datagenerazioneStr = datagenerazione.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture); //TODO: mese in formato testo, gestire locale
                //Dati Fondo
                row = 3;
                string classe = excelHelper.GetValue(mainSheetname, classeCol, row.ToString());
                while (!string.IsNullOrEmpty(classe))
                {
                    string currentIsin = excelHelper.GetValue(mainSheetname, isinCol, row.ToString());
                    SortedDictionary<string,string> performances = new SortedDictionary<string, string>();
                    isinPerformanceAnnoMap.TryGetValue(currentIsin, out performances);

                    m.KIIDData item = new m.KIIDData(
                        classe,
                        currentIsin,
                        excelHelper.GetValue(mainSheetname, classeDiRischioCol, row.ToString()),
                        excelHelper.GetValue(mainSheetname, testo1Col, row.ToString()),
                        excelHelper.GetValue(mainSheetname, testo2Col, row.ToString()),
                        excelHelper.GetValue(mainSheetname, testo3Col, row.ToString()),
                        excelHelper.GetValue(mainSheetname, spesesottoscrizioneCol, row.ToString()),
                        excelHelper.GetValue(mainSheetname, speserimborsoCol, row.ToString()),
                        excelHelper.GetValue(mainSheetname, spesecorrentiCol, row.ToString()),
                        excelHelper.GetValue(mainSheetname, spesediconversioneCol, row.ToString()),
                        excelHelper.GetValue(mainSheetname, commissioniRendimentoCol, row.ToString()),
                        excelHelper.GetValue(mainSheetname, informazionipraticheCol, row.ToString()),
                        datagenerazioneStr,
                        performances
                        );
                    result.Add(item);
                    row++;
                    classe = excelHelper.GetValue(mainSheetname, classeCol, row.ToString());
                    
                }

            }
            return result;
        }

        public void generateOutput(m.KIIDData data)
        {
            //Nome file--> nome fondo desunto dal template
            string templateName = template.Split('\\').LastOrDefault().Split('.').ElementAt(0);
            string outputFileName = outputfolder + "\\" + templateName + " Fund - KIID " + data.Classe + " " + datagenerazione.ToString("dd MM yyyy", CultureInfo.InvariantCulture) + " " + language + ".docx";
            using (KIIDWordHelper wordHelper = new KIIDWordHelper(template, outputFileName))
            {
                wordHelper.replaceText("@CLASSE@", data.Classe);
                wordHelper.replaceText("@ISIN@", data.Isin);
                wordHelper.replaceText("@SPESEDISOTTOSCRIZIONE@", string.Format("{0} %", data.SpeseSottoscrizione));
                wordHelper.replaceText("@TESTO1@", data.Testo1);
                wordHelper.InsertProfiloRischio(data.ClasseDiRischio);
                wordHelper.EditPerformanceTable(data.Performances);

            }
            using (WordHelper wordHelper = new WordHelper(outputFileName))
            {
                wordHelper.SaveAsPDF();
            }


        }

    }
}
