using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO.Packaging;
using System.Linq;
using System.Text;

namespace it.valuelab.hedgeinvest.helpers
{
    class ExcelHelper : IDisposable
    {

        private ExcelDocumentData excelData;

        public ExcelHelper(String filename)
        {
            excelData = new ExcelDocumentData(SpreadsheetDocument.Open(filename, false));

        }

        public String GetValue(String sheet, String col, String row)
        {
            WorksheetPart currentSheet = excelData.GetWorksheetPartByName(sheet);
            String reference = col + row;

            Cell currentCell = currentSheet.Worksheet.Descendants<Cell>().
                     Where(c => c.CellReference == reference).FirstOrDefault();

            string value = "";
            if (currentCell != null)
            {
                if (currentCell.CellValue != null)
                {
                    value = currentCell.CellValue.InnerText;
                }
                else
                {
                    value = currentCell.InnerText;
                }
                if ((currentCell.DataType != null) && (currentCell.DataType == CellValues.SharedString))
                {
                    value = excelData.SharedStringTablePart.SharedStringTable
                    .ChildElements[Int32.Parse(value)]
                    .InnerText;
                }

            }
            return value;
        }

        /// <summary>
        /// This event scans through the file to check if there is any files embedded in it.
        /// If there is any, it will add the name of the file in the checked list box
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void findEmbed(string fileName)
        {
            string embeddingPartString;
            Package pkg = Package.Open(fileName);
            StringBuilder embeddedFiles = new StringBuilder();

            System.IO.FileInfo fi = new System.IO.FileInfo(fileName);

            string extension = fi.Extension.ToLower();

            if ((extension == ".docx") || (extension == ".dotx") || (extension == ".docm") || (extension == ".dotm"))
            {
                embeddingPartString = "/word/embeddings/";
            }
            else if ((extension == ".xlsx") || (extension == ".xlsm") || (extension == ".xltx") || (extension == ".xltm"))
            {
                embeddingPartString = "/excel/embeddings/";
            }
            else
            {
                embeddingPartString = "/ppt/embeddings/";
            }

            // Get the embedded files names.
            foreach (PackagePart pkgPart in pkg.GetParts())
            {
                if (pkgPart.Uri.ToString().StartsWith(embeddingPartString))
                {
                    string fileName1 = pkgPart.Uri.ToString().Remove(0, embeddingPartString.Length);
                    embeddedFiles.Append(fileName1);
                }
            }
            pkg.Close();
            
        }

        public void Dispose()
        {
            excelData.Dispose();
        }
    }


    internal class ExcelDocumentData
    {
        private Dictionary<string, WorksheetPart> worksheets = new Dictionary<string, WorksheetPart>();

        private SpreadsheetDocument _document;

        public ExcelDocumentData(SpreadsheetDocument document)
        {
            _document = document;
            foreach (Sheet s in _document.WorkbookPart.Workbook.Descendants<Sheet>())
            {
                worksheets.Add(s.Name, (WorksheetPart)_document.WorkbookPart.GetPartById(s.Id));
            }

        }

        public WorksheetPart GetWorksheetPartByName(string name)
        {
            WorksheetPart res = null;
            if (worksheets.TryGetValue(name, out res))
            {
                return res;
            }
            else
            {
                throw new KeyNotFoundException("Sheet " + name + " non trovato");
            }
        }

        public SharedStringTablePart SharedStringTablePart
        {
            get
            {
                return _document.WorkbookPart.SharedStringTablePart;
            }
        }

        public void Dispose()
        {
            _document.Dispose();
        }
    }
}
