using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DotNetExcelToolsSurvey.Library
{
    public class OpenXMLSAXLib : IExcelOperations
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

                }
            }
            return columns;
        }
    }
}
