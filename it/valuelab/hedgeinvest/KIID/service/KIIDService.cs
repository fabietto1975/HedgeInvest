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
using System.Threading;
using Common.Logging;
using System.ComponentModel;

namespace it.valuelab.hedgeinvest.KIID.service
{
    public class KIIDService : INotifyPropertyChanged
    {
        private static readonly ILog Log = LogManager.GetLogger<KIIDService>();
        private string template;
        private string datafile;
        private string outputfolder;
        private string language;
        private string country;
        private DateTime datagenerazione;
        private CultureInfo cultureInfo;

        public event PropertyChangedEventHandler PropertyChanged;

        private int TotNumeroDocumenti = 0;
        private int CurrentNumeroDocumenti = 0;

        private double _progress;
        public double progress
        {
            get { return _progress; }
            set { _progress = value; }
        }
        private string _error;
        public string error
        {
            get { return _error; }
            set { _error = value; }
        }



        public KIIDService(string _template, string _datafile, string _outputfolder, string _language, DateTime _datagenerazione)
        {
            template = _template;
            datafile = _datafile;
            outputfolder = _outputfolder;
            datagenerazione = _datagenerazione;
            string[] splitLanguage = _language.Split('_');
            language = splitLanguage[0];
            country = splitLanguage[1];
            cultureInfo = new CultureInfo(language);
        }

        private List<m.KIIDData> ReadFundsData()

        {
            Log.Info("Inizio generazione documenti ");
            const string mainSheetname = "DATI KIID";
            const string performanceSheetname = "PERFORMANCE";
            const int MAX_EMPTY_ROWS = 3;
            List<m.KIIDData> result = new List<m.KIIDData>();
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en");
            Log.Info("Inizio lettura dati input");
            using (ExcelHelper excelHelper = new ExcelHelper(datafile))
            {

                //Performance
                Log.Info("Inizio lettura dati input - Performance");

                int row = 2;
                Dictionary<string, SortedDictionary<string,string>> isinPerformanceAnnoMap = new Dictionary<string, SortedDictionary<string, string>>() ;
                string isin = excelHelper.GetValue(performanceSheetname, "B", row.ToString());
                int emptyRows = 0;

                while (emptyRows<= MAX_EMPTY_ROWS)
                {
                    string anno= "0", dato = "0";
                    if (string.IsNullOrEmpty(isin))
                    {
                        emptyRows++;
                    }
                    else
                    {
                        emptyRows = 0;
                        anno = excelHelper.GetValue(performanceSheetname, "C", row.ToString());
                        dato = excelHelper.GetValue(performanceSheetname, "D", row.ToString());
                        if (string.IsNullOrEmpty(dato))
                        {
                            dato = "0";
                        } else
                        {
                            dato = (Convert.ToDouble(dato)/100) .ToString();
                        }
                        SortedDictionary<string, string> isinPerformanceAnno;
                        if (!isinPerformanceAnnoMap.TryGetValue(isin, out isinPerformanceAnno))
                        {
                            isinPerformanceAnno = new SortedDictionary<string, string>();
                        }
                        isinPerformanceAnno[anno] = dato;
                        isinPerformanceAnnoMap[isin] = isinPerformanceAnno;
                    }
                    row++;
                    isin = excelHelper.GetValue(performanceSheetname, "B", row.ToString());

                }
                Log.Info("Termine lettura dati input - Performance");
                Log.Info("Inizio lettura dati input - Dati fondi");

                string suffix = "";
                if (!language.Equals("it-IT"))
                {
                    suffix += " - " + language.Split('-')[0].ToUpper(); 
                }
                Log.Info("Dati fondi: suffisso " + suffix);
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
                string datagenerazioneStr = datagenerazione.ToString("dd MMMM yyyy", cultureInfo );
                Log.Debug("testo1Col  " + testo1Col);
                Log.Debug("testo2Col  " + testo2Col);
                Log.Debug("testo3Col  " + testo3Col);
                Log.Debug("commissioniRendimentoCol  " + commissioniRendimentoCol);
                Log.Debug("informazionipraticheCol  " + informazionipraticheCol);
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
                        (Convert.ToDouble(excelHelper.GetValue(mainSheetname, spesesottoscrizioneCol, row.ToString()))*100).ToString(),
                        (Convert.ToDouble(excelHelper.GetValue(mainSheetname, speserimborsoCol, row.ToString())) * 100).ToString(),
                        (Convert.ToDouble(excelHelper.GetValue(mainSheetname, spesecorrentiCol, row.ToString())) * 100).ToString(),
                        (Convert.ToDouble(excelHelper.GetValue(mainSheetname, spesediconversioneCol, row.ToString())) * 100).ToString(),
                        excelHelper.GetValue(mainSheetname, commissioniRendimentoCol, row.ToString()),
                        excelHelper.GetValue(mainSheetname, informazionipraticheCol, row.ToString()),
                        datagenerazioneStr,
                        performances
                        );
                    result.Add(item);
                    row++;
                    classe = excelHelper.GetValue(mainSheetname, classeCol, row.ToString());
                    
                }
                Log.Info("Termine lettura dati input - Dati fondi");

            }
            TotNumeroDocumenti = result.Count();
            return result;
        }

        private void GenerateOutput(m.KIIDData data)
        {
            //Nome file--> nome fondo desunto dal template
            string templateName = template.Split('\\').LastOrDefault().Split('.').ElementAt(0);
            
            string outputFileName = outputfolder + "\\"  + "KIID_" + templateName + "_" + data.Classe  + "_" + data.Isin+ "_" + language.Split('-')[1]+ "_" + country + "_" + datagenerazione.ToString("dd MM yyyy", CultureInfo.InvariantCulture) + ".docx";
            Log.Info("Inizio generazione documento " + (CurrentNumeroDocumenti + 1) + " di " + TotNumeroDocumenti + " " + outputFileName );
            using (KIIDWordHelper wordHelper = new KIIDWordHelper(template, outputFileName))
            {
                wordHelper.ReplaceText("@CLASSE@", data.Classe);
                wordHelper.ReplaceText("@ISIN@", data.Isin);
                wordHelper.ReplaceText("@TESTO1@", data.Testo1, "FORMATTED");
                wordHelper.ReplaceText("@CLASSEDIRISCHIO@", data.ClasseDiRischio);
                wordHelper.ReplaceText("@SPESEDISOTTOSCRIZIONE@", string.Format("{0}%", data.SpeseSottoscrizione));
                wordHelper.ReplaceText("@SPESEDIRIMBORSO@", string.Format("{0}%", data.SpeseDiRimborso));
                wordHelper.ReplaceText("@SPESEDICONVERSIONE@", string.Format("{0}%", data.SpeseDiConversione));
                wordHelper.ReplaceText("@SPESECORRENTI@", string.Format("{0}%", data.SpeseCorrenti));
                wordHelper.ReplaceText("@COMMISSIONIRENDIMENTO@", data.CommissioniRendimento);
                wordHelper.ReplaceText("@DATAGENERAZIONE@", data.DataGenerazione);
                wordHelper.ReplaceText("@INFORMAZIONIPRATICHE", data.InformazioniPratiche);
                wordHelper.InsertProfiloRischio(data.ClasseDiRischio);
                wordHelper.EditPerformanceChart(data.Performances);
                wordHelper.ReplaceText("@TESTO2@", data.Testo2);
                wordHelper.ReplaceText("@TESTO3@", data.Testo3);
            }
            using (WordHelper wordHelper = new WordHelper(outputFileName))
            {
                wordHelper.SaveAsPDF();
            }
            Log.Info("Generazione documento " + (CurrentNumeroDocumenti+1) + " di " + TotNumeroDocumenti + " " + outputFileName + " terminata");

            CurrentNumeroDocumenti++;
            progress = (double)CurrentNumeroDocumenti / TotNumeroDocumenti;
            NotifyPropertyChanged("progress");

        }


        public void GenerateKIID()
        {
            try
            {
                List<m.KIIDData> kiidDataList = this.ReadFundsData();
                foreach (m.KIIDData kiiddata in kiidDataList)
                {
                    GenerateOutput(kiiddata);
                }

            }
            catch (Exception e)
            {
                NotifyPropertyChanged("error");
                error = "Errore nell'esecuzione della procedura: " + e.Message;
            }

        }

        private void NotifyPropertyChanged(string property)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(property));
            }
        }


    }
}
