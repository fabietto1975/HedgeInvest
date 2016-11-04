using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace it.valuelab.hedgeinvest.KIID.model
{

    public sealed class KIIDData
    {

        private String _template;
        public string Template { get { return _template; }  }
        private String _classe;
        public String Classe { get { return _classe; } }
        private string _isin;
        public String Isin { get { return _isin; } }
        private String _classeDiRischio;
        public string ClasseDiRischio { get { return _classeDiRischio; } }
        private String _testo1;
        public String Testo1 { get { return _testo1; } }
        private String _testo2;
        public String Testo2 { get { return _testo2; } }
        private String _testo3;
        public String Testo3 { get { return _testo3; } }
        private String _speseSottoscrizione;
        public String SpeseSottoscrizione { get { return _speseSottoscrizione; } }
        private string _dataKIID;
        public string DataKIID { get { return _dataKIID; } }
        private String _speseDiRimborso;
        public String SpeseDiRimborso { get { return _speseDiRimborso; } }
        private String _speseCorrenti;
        public String SpeseCorrenti { get { return _speseCorrenti; } }
        private string _speseDiConversione;
        public string SpeseDiConversione { get { return _speseDiConversione; } }
        private String _commissioniRendimento;
        public String CommissioniRendimento { get { return _commissioniRendimento; } }
        public String _informazioniPratiche;
        public String InformazioniPratiche { get { return _informazioniPratiche; } }
        private SortedDictionary<string, string> _performances;
        public SortedDictionary<string, string> Performances { get { return _performances; } }

        public KIIDData(String template, String classe, String isin, String classeDiRischio, String testo1, String testo2, String testo3, string datakiid, string speseSottoscrizione,
                String speseDiRimborso, String speseCorrenti, String speseDiConversione, String commissioniRendimento, String informazioniPratiche, SortedDictionary<string, string> performances)
        {

            _template = template;
            _classe = classe;
            _isin = isin;
            _classeDiRischio = classeDiRischio;
            _testo1 = testo1;
            _testo2 = testo2;
            _testo3 = testo3;
            _dataKIID = datakiid;
            _speseSottoscrizione = speseSottoscrizione;
            _speseDiRimborso = speseDiRimborso;
            _speseCorrenti = speseCorrenti;
            _speseDiConversione = speseDiConversione;
            _commissioniRendimento = commissioniRendimento;
            _informazioniPratiche = informazioniPratiche;
            _performances = performances;
        }
    }
}
