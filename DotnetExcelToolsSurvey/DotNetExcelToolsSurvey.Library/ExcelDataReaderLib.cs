using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DotNetExcelToolsSurvey.Library
{
    public class ExcelDataReaderLib : IExcelOperations
    {
        public List<string> GetSheetNames(string path)
        {
            List<string> sheetNames = new List<string>();

            using (var stream = File.Open(path, FileMode.Open, FileAccess.Read))
            {

                // Auto-detect format, supports:
                //  - Binary Excel files (2.0-2003 format; *.xls)
                //  - OpenXml Excel files (2007 format; *.xlsx)
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {

                    // Choose one of either 1 or 2:

                    // 1. Use the reader methods
                    do
                    {
                        sheetNames.Add(reader.Name);
                        //while (reader.Read())
                        //{
                        //    // reader.GetDouble(0);
                        //}
                    } while (reader.NextResult());
                }
            }


            return sheetNames;
        }
        public List<List<string>> GetFirstRow(string path)
        {
            List<List<string>> columns = new List<List<string>>();

            using (var stream = File.Open(path, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {

                    // Choose one of either 1 or 2:

                    // 1. Use the reader methods
                    
                    do
                    {
                        List<string> columnValues = new List<string>();
                        reader.Read();
                        columnValues.Add(reader.Name);
                        for (int i = 0; i < reader.FieldCount; i++)
                        {
                            columnValues.Add(reader.GetString(i));
                        }

                        columns.Add(columnValues);
                    } while (reader.NextResult());
                }
            }
            return columns;
        }
    }
}
