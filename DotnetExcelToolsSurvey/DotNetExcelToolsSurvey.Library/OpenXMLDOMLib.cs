using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.WindowsAPICodePack.Dialogs;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace DotNetExcelToolsSurvey.Library
{
    public class OpenXMLDOMLib : IExcelOperations
    {
        public List<string> GetSheetNames(string path)
        {
            List<string> sheetNames = new List<string>();
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(path, false))
            {
                WorkbookPart wbPart = document.WorkbookPart;
                Sheets theSheets = wbPart.Workbook.Sheets;
                foreach (Sheet sheet in theSheets)
                {
                    sheetNames.Add(sheet.Name);
                }
            }
            return sheetNames;
        }

        public List<List<string>> GetFirstRow(string path)
        {
            List<List<string>> columns = new List<List<string>>();
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(path, false))
            {
                WorkbookPart wbPart = document.WorkbookPart;
                Sheets theSheets = wbPart.Workbook.Sheets;
                foreach (Sheet sheet in theSheets)
                {
                    List<string> columnValues = new List<string>();
                    WorksheetPart wsPart = (WorksheetPart)wbPart.GetPartById(sheet.Id);
                    SheetData sheetData = wsPart.Worksheet.Elements<SheetData>().First();
                    columnValues.Add(sheet.Name);
                    Row r = sheetData.Elements<Row>().First();
                    foreach (Cell c in r.Elements<Cell>())
                    {
                        columnValues.Add(c.CellValue.Text);
                    }
                    columns.Add(columnValues);
                }
            }
            return columns;
        }
    }
}
