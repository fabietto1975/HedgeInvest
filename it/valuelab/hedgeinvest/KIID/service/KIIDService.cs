using m = it.valuelab.hedgeinvest.KIID.model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using it.valuelab.hedgeinvest.helpers;

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
                string isin = excelHelper.GetValue(performanceSheetname, "A", row.ToString());
                while (!isin.Equals(""))
                {

                    string anno = excelHelper.GetValue(performanceSheetname, "B", row.ToString());
                    string dato = excelHelper.GetValue(performanceSheetname, "C", row.ToString());

                    SortedDictionary<string, string> isinPerformanceAnno; 
                    if (!isinPerformanceAnnoMap.TryGetValue(isin, out isinPerformanceAnno))
                    {
                        isinPerformanceAnno = new SortedDictionary<string, string>();
                    }
                    isinPerformanceAnno[anno] = dato;

                    isinPerformanceAnnoMap[isin] =  isinPerformanceAnno;
                    row++;
                    isin = excelHelper.GetValue(performanceSheetname, "A", row.ToString());
                }

                //Dati Fondo
                row = 2;
                string classe = excelHelper.GetValue(mainSheetname, "A", row.ToString());
                while (!classe.Equals(""))
                {
                    string currentIsin = excelHelper.GetValue(mainSheetname, "B", row.ToString());
                    System.Diagnostics.Debug.WriteLine("currentIsin:" + currentIsin);
                    SortedDictionary<string,string> performances = new SortedDictionary<string, string>();
                    isinPerformanceAnnoMap.TryGetValue(isin, out performances);
                    m.KIIDData item = new m.KIIDData(
                        classe,
                        currentIsin,
                        excelHelper.GetValue(mainSheetname, "B", row.ToString()),
                        excelHelper.GetValue(mainSheetname, "D", row.ToString()),
                        excelHelper.GetValue(mainSheetname, "E", row.ToString()),
                        excelHelper.GetValue(mainSheetname, "F", row.ToString()),
                        excelHelper.GetValue(mainSheetname, "H", row.ToString()),
                        excelHelper.GetValue(mainSheetname, "I", row.ToString()),
                        excelHelper.GetValue(mainSheetname, "J", row.ToString()),
                        excelHelper.GetValue(mainSheetname, "K", row.ToString()),
                        excelHelper.GetValue(mainSheetname, "L", row.ToString()),
                        excelHelper.GetValue(mainSheetname, "M", row.ToString()),
                        excelHelper.GetValue(mainSheetname, "N", row.ToString()),
                        performances
                        );
                    result.Add(item);
                    classe = excelHelper.GetValue(mainSheetname, "A", row.ToString());
                    row++;
                }

            }
            return result;
        }

        public void generateOutput(m.KIIDData data)
        {
            const string outPath = @"D:\LAVORO\PROGETTI\HEDGEINVEST\KKID\OUT"; //TODO: esternalizzare property
            const string templatePath = @"D:\LAVORO\PROGETTI\HEDGEINVEST\KKID\TEMPLATE"; //TODO: esternalizzare property

            string inputFileName = templatePath+"\\" +data.Classe + ".docx"; ;
            string outputFileName = outPath + "\\" + data.Classe + "_" + data.Isin+ ".docx";
            System.Diagnostics.Debug.WriteLine(inputFileName);
            System.Diagnostics.Debug.WriteLine(outputFileName);
            using (WordHelper wordHelper = new WordHelper(inputFileName, outputFileName))
            {
                wordHelper.replaceText("@CLASSE@", data.Classe);
                wordHelper.replaceText("@ISIN@", data.Isin);
            }

        }

    }
}
