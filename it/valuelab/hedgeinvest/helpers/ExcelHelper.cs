using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace it.valuelab.hedgeinvest.helpers
{
    class ExcelHelper : IDisposable
    {

        private ExcelDocumentData excelData;

        public ExcelHelper(String filename)
        {
            excelData = new ExcelDocumentData(SpreadsheetDocument.Open(filename, false));

        }

        public int GetSheetRowCount(string sheet)

        {
            WorksheetPart currentSheet = excelData.GetWorksheetPartByName(sheet);
            SheetData sd = currentSheet.Worksheet.Elements<SheetData>().FirstOrDefault();
            return sd.Elements<Row>().Count();
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
                value = GetCellText(currentCell);

            }
            return value;
        }

        public Cell GetCellByContent(string sheet, string content, int rowindex)
        {
            WorksheetPart currentSheet = excelData.GetWorksheetPartByName(sheet);
            SheetData sd = currentSheet.Worksheet.Elements<SheetData>().FirstOrDefault();
            Row row = sd.Elements<Row>().ElementAt(rowindex);
            Cell result = row.Elements<Cell>().Where(
                c =>  content.Equals(GetCellText(c))
                ).FirstOrDefault();
            return result;
        }

        public string GetCellColumn(Cell c)
        {
            if (c !=null && c.CellReference != null)
            {
                return Regex.Replace(c.CellReference, "[0-9]", ""); 
            }
            else
            {
                return null;
            }
        }

        private string GetCellText(Cell c)
        {
            string value;
            if (c.CellValue != null)
            {
                value = c.CellValue.InnerText;
            }
            else
            {
                value = c.InnerText;
            }
            if ((c.DataType != null) && (c.DataType == CellValues.SharedString))
            {
                value = excelData.SharedStringTablePart.SharedStringTable
                .ChildElements[Int32.Parse(value)]
                .InnerText;
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
