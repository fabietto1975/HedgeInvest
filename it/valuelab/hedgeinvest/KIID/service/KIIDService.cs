using m = it.valuelab.hedgeinvest.KIID.model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using it.valuelab.hedgeinvest.helpers;
using KIID.it.valuelab.hedgeinvest.KIID.helpers;

namespace it.valuelab.hedgeinvest.KIID.service
{
    public class KIIDService
    {
        public List<m.KIIDData> readFundsData()

        {
            string inputFileName = @"D:\LAVORO\PROGETTI\HEDGEINVEST\KKID\INPUT\DATIKIDD.XLSX"; //TODO: esternalizzare property
            const string mainSheetname = "DATI KIID";
            const string performanceSheetname = "PERFORMANCE";
            List<m.KIIDData> result = new List<m.KIIDData>();
            using (ExcelHelper excelHelper = new ExcelHelper(inputFileName))
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

                //Dati Fondo
                row = 3;
                string classe = excelHelper.GetValue(mainSheetname, "C", row.ToString());
                while (!string.IsNullOrEmpty(classe))
                {
                    string currentIsin = excelHelper.GetValue(mainSheetname, "D", row.ToString());
                    SortedDictionary<string,string> performances = new SortedDictionary<string, string>();
                    isinPerformanceAnnoMap.TryGetValue(currentIsin, out performances);
                    m.KIIDData item = new m.KIIDData(
                        excelHelper.GetValue(mainSheetname, "B", row.ToString()),
                        classe,
                        currentIsin,
                        excelHelper.GetValue(mainSheetname, "E", row.ToString()),
                        excelHelper.GetValue(mainSheetname, "F", row.ToString()),
                        excelHelper.GetValue(mainSheetname, "G", row.ToString()),
                        excelHelper.GetValue(mainSheetname, "H", row.ToString()),
                        excelHelper.GetValue(mainSheetname, "J", row.ToString()),
                        excelHelper.GetValue(mainSheetname, "K", row.ToString()),
                        excelHelper.GetValue(mainSheetname, "L", row.ToString()),
                        excelHelper.GetValue(mainSheetname, "M", row.ToString()),
                        excelHelper.GetValue(mainSheetname, "N", row.ToString()),
                        excelHelper.GetValue(mainSheetname, "O", row.ToString()),
                        excelHelper.GetValue(mainSheetname, "P", row.ToString()),
                        performances
                        );
                    result.Add(item);
                    row++;
                    classe = excelHelper.GetValue(mainSheetname, "C", row.ToString());
                    
                }

            }
            return result;
        }

        public void generateOutput(m.KIIDData data)
        {
            const string outPath = @"D:\LAVORO\PROGETTI\HEDGEINVEST\KKID\OUT"; //TODO: esternalizzare property
            const string templatePath = @"D:\LAVORO\PROGETTI\HEDGEINVEST\KKID\TEMPLATE"; //TODO: esternalizzare property

            string inputFileName = templatePath+"\\" +data.Template + ".docx"; ;
            string outputFileName = outPath + "\\" + data.Template + "_" + data.Isin+ ".docx";
            using (KIIDWordHelper wordHelper = new KIIDWordHelper(inputFileName, outputFileName))
            {
                wordHelper.replaceText("@CLASSE@", data.Classe);
                wordHelper.replaceText("@ISIN@", data.Isin);
                wordHelper.replaceText("@SPESEDISOTTOSCRIZIONE@", string.Format("{0} %", data.SpeseSottoscrizione));
                //wordHelper.replaceText("@TESTO1@", data.Testo1);
                System.Diagnostics.Debug.WriteLine(data.Testo1);
                wordHelper.replaceText("@TESTO1@", "\t\u2022 Riga 1 \u000a\u000d\u2022 Riga 2");
                //wordHelper.replaceText("@GRAFICO@", "AA");
                wordHelper.InsertProfiloRischio(data.ClasseDiRischio);
                if (data.Performances != null)
                {
                    wordHelper.InsertPerformanceTable(data.Performances);
                }
            }

        }

    }
}
