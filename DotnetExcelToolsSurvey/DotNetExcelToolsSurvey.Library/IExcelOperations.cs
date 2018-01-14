using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DotNetExcelToolsSurvey.Library
{
    public interface IExcelOperations
    {
        List<string> GetSheetNames(string path);
        List<List<string>> GetFirstRow(string path);
        //void UpdateSheetNames(string path, List<string> newNames);
        //void UpdateFirstRow(string path, List<string> newNames);
    }
}
