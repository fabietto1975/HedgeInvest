using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace it.valuelab.hedgeinvest.helpers
{
    class WordHelper : IDisposable
    {

        protected WordprocessingDocument _document;
        protected WordprocessingDocument Document { get { return _document; } }
        

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
            docText = regexText.Replace(docText, newtext);
            using (StreamWriter sw = new StreamWriter(Document.MainDocumentPart.GetStream(FileMode.Create)))
            {
                sw.Write(docText);
            }
        }



        public void Dispose()
        {
            Document.Dispose();
        }
    }
}
