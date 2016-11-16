using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using s = it.valuelab.hedgeinvest.KIID.service;
using System.Diagnostics;
using System.IO;
using d = DocumentFormat.OpenXml.Drawing;
using dc = DocumentFormat.OpenXml.Drawing.Charts;
using dw = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using m = it.valuelab.hedgeinvest.KIID.model;
using System.Collections.Generic;
using System;
using Common.Logging;

namespace it.valuelab.hedgeinvest.KIID
{
       
    class CreateKIIDOpenXML
    {
        //private static ILog log = LogManager.GetLogger<CreateKIIDOpenXML>();
        private static readonly ILog Log = LogManager.GetLogger(typeof(CreateKIIDOpenXML));
        static void Main()
        {

            s.KIIDService service = new s.KIIDService(@"D:\LAVORO\PROGETTI\HEDGEINVEST\KIID\TEMPLATE\HICU_IT.docx", @"D:\LAVORO\PROGETTI\HEDGEINVEST\KIID\INPUT\DATIKIID.XLSX",
                    @"D:\LAVORO\PROGETTI\HEDGEINVEST\KIID\OUT","it-IT_IT", DateTime.Now);

            Log.Debug("Inizio creazione documenti KIID");
            List<m.KIIDData> kiidDataList = service.readFundsData();
            foreach (m.KIIDData kiiddata in kiidDataList)
            {
                service.generateOutput(kiiddata);
            }
            Log.Info("Creazione documenti KIID terminata");


        }
    }
}
