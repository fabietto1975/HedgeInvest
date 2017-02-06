using Common.Logging;
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
using wp = DocumentFormat.OpenXml.Wordprocessing;
using System.Runtime.Serialization;

namespace it.valuelab.hedgeinvest.helpers
{
    class WordHelper : IDisposable
    {
        protected static readonly ILog Log = LogManager.GetLogger<WordHelper>();

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

        public DocumentFormat.OpenXml.Wordprocessing.Table FindByCaption(string caption)
        {

            Body body = _document.MainDocumentPart.Document.Body;
            IEnumerable<TableProperties> tableProperties = body.Descendants<TableProperties>().Where(tp => tp.TableCaption != null);

            foreach (TableProperties tProp in tableProperties)
            {
                if (tProp.TableCaption.Val==caption) // see comment, this is actually StringValue
                {
                    // do something for table with myCaption
                    DocumentFormat.OpenXml.Wordprocessing.Table table = (DocumentFormat.OpenXml.Wordprocessing.Table)tProp.Parent;
                    return table;
                }
            }
            throw new TableNotFoundException("Impossible trovare la tabella con caption " + caption); 

        }



        public void RemoveRowByContent(string texttofind)
        {
            Body body = _document.MainDocumentPart.Document.Body;
            Text item = body.Descendants<Text>().Where(t => t.Text == texttofind).FirstOrDefault();
            OpenXmlElement ancestor;
            OpenXmlElement currItem = item;
            while (true && item !=null)
            {
                ancestor = currItem.Parent;
              
                if (ancestor == null)
                {
                    break;
                }
                else if (ancestor is TableRow)
                {
                    ancestor.Remove();
                    Log.Debug("REMOVED ROW");
                    break;
                }
          

                currItem = ancestor;
            }
            return;
        }
        public void ReplaceText(string oldtext, string newtext, string mode="PLAIN")
        {
            string docText = null;
            Body body =_document.MainDocumentPart.Document.Body;
            Text item = body.Descendants<Text>().Where(t => t.Text == oldtext).FirstOrDefault();
            if (item != null)
            {
                if (mode.Equals("PLAIN"))
                {
                    item.Text = newtext;

                }
                else
                {
                    /*
                    Run currentRun = (Run) item.Parent;
                    currentRun.RemoveAllChildren<Text>();
                    */

                    wp.Paragraph currentParagraph = (wp.Paragraph) item.Parent.Parent;
                    Run currentRun = (Run)item.Parent;
                    currentRun.RemoveAllChildren<Text>();
                    
                    //Gestione a capo
                    string[] splitted = newtext.Split('\n');
                    if (splitted.Length > 1)
                    {
                        item.Text = splitted[0];
                        Text currentTextBlock = item;
                        /*
                        Text newTextBlock = new Text(splitted[0]);
                        currentRun.Append(newTextBlock);
                        */

                        Text newTextBlock = new Text(splitted[0]);
                        currentRun.Append(newTextBlock);
                        for (int idx = 1; idx < splitted.Length; idx++)
                        {

                            Run newRun = new Run();
                            newRun.RunProperties = (RunProperties)currentRun.RunProperties.CloneNode(true);
                            newTextBlock = new Text(splitted[idx]);
                            newRun.Append(newTextBlock);
                            currentRun = newRun;
                            wp.Paragraph newParagraph = new wp.Paragraph(newRun);
                            newParagraph.ParagraphProperties = (ParagraphProperties)currentParagraph.ParagraphProperties.CloneNode(true);
                            currentParagraph.InsertAfterSelf(newParagraph);
                            currentParagraph = newParagraph;
                            /*
                            wp.Break b = new wp.Break();
                            currentTextBlock.InsertAfterSelf(b);
                            Text newTextBlock = new Text(splitted[idx]);
                            b.InsertAfterSelf(newTextBlock);
                            currentTextBlock = newTextBlock;
                            */

                            /*
                            Run newRun = new Run();
                            newRun.RunProperties = (RunProperties) currentRun.RunProperties.CloneNode(true);
                            newTextBlock = new Text(splitted[idx]);
                            newRun.AppendChild(newTextBlock);
                            currentRun.InsertAfterSelf(newRun);
                            newTextBlock.InsertAfterSelf(new DocumentFormat.OpenXml.Wordprocessing.Break());
                            currentRun = newRun;
                            */
                        }

                    }
                }

            }
            using (StreamWriter sw = new StreamWriter(Document.MainDocumentPart.GetStream(FileMode.Create)))
            {
                sw.Write(docText);
            }
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

        [Serializable]
        private class TableNotFoundException : Exception
        {
            public TableNotFoundException()
            {
            }

            public TableNotFoundException(string message) : base(message)
            {
            }

            public TableNotFoundException(string message, Exception innerException) : base(message, innerException)
            {
            }

            protected TableNotFoundException(SerializationInfo info, StreamingContext context) : base(info, context)
            {
            }
        }
    }
}
