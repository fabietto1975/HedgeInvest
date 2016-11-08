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

        public Cell GetCellByContent(string sheet, string content, int rowindex)
        {
            WorksheetPart currentSheet = excelData.GetWorksheetPartByName(sheet);
            SheetData sd = currentSheet.Worksheet.Elements<SheetData>().FirstOrDefault();
            Row row = sd.Elements<Row>().ElementAt(rowindex);
            foreach(Cell c in row.Elements<Cell>())
            {
                if ((c.DataType != null) && (c.DataType == CellValues.SharedString))
                    System.Diagnostics.Debug.WriteLine(excelData.SharedStringTablePart.SharedStringTable
                    .ChildElements[int.Parse(c.InnerText)]
                    .InnerText);
            }
            Cell result = row.Elements<Cell>().Where(
                c => c.CellValue.Text == content
                ).FirstOrDefault();
            return result;
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
                throw new KeyNotFoundException("Sheet "+ name +" non trovato");
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
