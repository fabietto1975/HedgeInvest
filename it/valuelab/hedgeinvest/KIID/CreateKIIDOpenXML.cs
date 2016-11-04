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

namespace it.valuelab.hedgeinvest.KIID
{
    class CreateKIIDOpenXML
    {
        static void Main()
        {
            s.KIIDService service = new s.KIIDService();
            List<m.KIIDData> kiidDataList = service.readFundsData();
            //Prova commit
            foreach (m.KIIDData kiiddata in kiidDataList)
            {
                service.generateOutput(kiiddata);
            }
            

        }
    }
}
