using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Office.Interop.Word;
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

        protected string _filename;
        protected string _outName;
        protected WordprocessingDocument _document;
        protected WordprocessingDocument Document { get { return _document; } }
        // test modifica

        public WordHelper(String filename)
        {
            _filename = filename;
            _document = WordprocessingDocument.Open(filename, false);
        }

        public WordHelper(String filename, String outName)
        {
            File.Copy(filename, outName, true);
            _document = WordprocessingDocument.Open(outName, true);
            _filename = filename;
            _outName = outName;
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

        public void SaveAsPDF()
        {
            Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
            object oMissing = System.Reflection.Missing.Value;
            // Cast as Object for word Open method

            Object filename = (Object)_filename;
            
            Microsoft.Office.Interop.Word.Document doc = word.Documents.Open(ref filename, ref oMissing,
                    ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing, ref oMissing, ref oMissing);
            doc.Activate();

            object outputFileName = _filename.Replace(".docx", ".pdf");
            object fileFormat = WdSaveFormat.wdFormatPDF;

                // Save document into PDF Format
            doc.SaveAs(ref outputFileName,
                ref fileFormat, ref oMissing, ref oMissing,
                ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                ref oMissing, ref oMissing, ref oMissing, ref oMissing);
            object saveChanges = WdSaveOptions.wdDoNotSaveChanges;
            ((_Document)doc).Close(ref saveChanges, ref oMissing, ref oMissing);
            doc = null;

            // word has to be cast to type _Application so that it will find
            // the correct Quit method.
            ((_Application)word).Quit(ref oMissing, ref oMissing, ref oMissing);
                        word = null;
        }

        public void Dispose()
        {
            Document.Dispose();
        }
    }
}
