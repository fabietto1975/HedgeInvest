using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace it.valuelab.hedgeinvest.helpers
{
    class WordHelper : IDisposable
    {

        protected WordprocessingDocument _document;
        protected WordprocessingDocument Document { get { return _document; } }
        // test modifica

        public WordHelper(String filename)
        {
            _document = WordprocessingDocument.Open(filename, false);
        }

        public WordHelper(String filename, String outName)
        {
            File.Copy(filename, outName, true);
            _document = WordprocessingDocument.Open(outName, true);
        }

        public void replaceText(string oldtext, string newtext)
        {
            string docText = null;

            using (StreamReader sr = new StreamReader(Document.MainDocumentPart.GetStream()))
            {
                docText = sr.ReadToEnd();
            }

            Regex regexText = new Regex(oldtext);
            newtext = formattaTesto(newtext);
            docText = regexText.Replace(docText, newtext);
            using (StreamWriter sw = new StreamWriter(Document.MainDocumentPart.GetStream(FileMode.Create)))
            {
                sw.Write(docText);
            }
        }
        private string formattaTesto(string newtext)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("<w:t>");
            sb.Append(newtext.Replace("\n","<w:br/>"));
            sb.Append("</w:t>");

            return sb.ToString();
        }



        public void Dispose()
        {
            Document.Dispose();
        }
    }
}
