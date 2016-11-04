using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;

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
        
        public void Dispose()
        {
            excelData.Dispose();
        }
    }


    internal class ExcelDocumentData
    {
        private Dictionary<string, WorksheetPart> worksheets = new Dictionary<string, WorksheetPart>();
        private SpreadsheetDocument Document { get; }

        public ExcelDocumentData(SpreadsheetDocument _document)
        {
            Document = _document;
            foreach (Sheet s in Document.WorkbookPart.Workbook.Descendants<Sheet>())
            {
                worksheets.Add(s.Name, (WorksheetPart)Document.WorkbookPart.GetPartById(s.Id));
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
                throw new KeyNotFoundException("Sheet "+ name +" non trovato");
            }
        }

        public SharedStringTablePart SharedStringTablePart
        {
            get 
            {
                return Document.WorkbookPart.SharedStringTablePart;
            }
        }

        public void Dispose()
        {
            Document.Dispose();
        }
    }
}
