using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace it.valuelab.hedgeinvest.KIID.model
{

    public sealed class KIIDData
    {
        public String Classe { get; }
        public String Isin { get; }
        public String ClasseDiRischio { get; }
        public String Testo1 { get; }
        public String Testo2 { get; }
        public String Testo3 { get; }
        public String SpeseSottoscrizione { get; }
        public string DataKIID { get; }
        public String SpeseDiRimborso { get; }
        public String SpeseCorrenti { get; }
        public String SpeseDiConversione { get; }
        public String CommissioniRendimento { get; }
        public String InformazioniPratiche { get; }
        public SortedDictionary<string, string> Performances { get; }

        public KIIDData(String classe, String isin, String classeDiRischio, String testo1, String testo2, String testo3, string datakiid, string speseSottoscrizione,
                String speseDiRimborso, String speseCorrenti, String speseDiConversione, String commissioniRendimento, String informazioniPratiche, SortedDictionary<string, string> performances)
        {
            
            Classe = classe;
            Isin = isin;
            ClasseDiRischio = classeDiRischio;
            Testo1 = testo1;
            Testo2 = testo2;
            Testo3 = testo3;
            DataKIID = datakiid;
            SpeseSottoscrizione = speseSottoscrizione;
            SpeseDiRimborso = speseDiRimborso;
            SpeseCorrenti = speseCorrenti;
            SpeseDiConversione = speseDiConversione;
            CommissioniRendimento = commissioniRendimento;
            InformazioniPratiche = informazioniPratiche;
            Performances = performances;
        }
    }
}
