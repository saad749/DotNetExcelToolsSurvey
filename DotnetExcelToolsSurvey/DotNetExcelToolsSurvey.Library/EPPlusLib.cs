using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DotNetExcelToolsSurvey.Library
{
    public class EPPlusLib : IExcelOperations
    {
        public List<string> GetSheetNames(string path)
        {
            List<string> sheetNames = new List<string>();



            return sheetNames;
        }
        public List<List<string>> GetFirstRow(string path)
        {
            List<List<string>> columns = new List<List<string>>();

            return columns;
        }
    }
}
