using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace it.valuelab.hedgeinvest.KIID.model
{

    public sealed class KIIDData
    {

        private String _classe;
        public String Classe { get { return _classe; } }
        private string _isin;
        public string Isin { get { return _isin; } }
        private string _classeDiRischio;
        public string ClasseDiRischio { get { return _classeDiRischio; } }
        private string _testo1;
        public string Testo1 { get { return _testo1; } }
        private string _testo2;
        public string Testo2 { get { return _testo2; } }
        private string _testo3;
        public string Testo3 { get { return _testo3; } }
        private string _speseSottoscrizione;
        public string SpeseSottoscrizione { get { return _speseSottoscrizione; } }
        private string _speseDiRimborso;
        public string SpeseDiRimborso { get { return _speseDiRimborso; } }
        private string _speseCorrenti;
        public string SpeseCorrenti { get { return _speseCorrenti; } }
        private string _speseDiConversione;
        public string SpeseDiConversione { get { return _speseDiConversione; } }
        private String _commissioniRendimento;
        public String CommissioniRendimento { get { return _commissioniRendimento; } }
        private string _informazioniPratiche;
        public string InformazioniPratiche { get { return _informazioniPratiche; } }
        private string _dataGenerazione;
        public string DataGenerazione { get { return _dataGenerazione; } }

        private SortedDictionary<string, string> _performances;
        public SortedDictionary<string, string> Performances { get { return _performances; } }

        public KIIDData(String classe, String isin, String classeDiRischio, String testo1, String testo2, String testo3, string speseSottoscrizione,
                String speseDiRimborso, String speseCorrenti, String speseDiConversione, String commissioniRendimento, String informazioniPratiche, String dataGenerazione, SortedDictionary<string, string> performances)
        {
            _classe = classe;
            _isin = isin;
            _classeDiRischio = classeDiRischio;
            _testo1 = testo1;
            _testo2 = testo2;
            _testo3 = testo3;
            _speseSottoscrizione = speseSottoscrizione;
            _speseDiRimborso = speseDiRimborso;
            _speseCorrenti = speseCorrenti;
            _speseDiConversione = speseDiConversione;
            _commissioniRendimento = commissioniRendimento;
            _informazioniPratiche = informazioniPratiche;
            _dataGenerazione = dataGenerazione;
            _performances = performances;
        }
    }
}
