using DotNetExcelToolsSurvey.Library;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace DotNetExcelToolsSurvey.Tests
{
    public class OpenXMLDomLibTests
    {
        [Fact]
        public void IsSheetCountCorrect()
        {
            string path = @"C:\Users\PHCCUser\Desktop\PMS Data\2017\DW_2017\LAB_TestsPerformed_2017\Lab Data Oct-2017_Sample-collected-orders_Tests.xlsx";

            IExcelOperations excelOps = new OpenXMLDOMLib();
            Assert.Equal(2, excelOps.GetSheetNames(path).Count);
        }

        [Fact]
        public void IsReadingAllSheetNames()
        {
            string path = @"C:\Users\PHCCUser\Desktop\PMS Data\2017\DW_2017\LAB_TestsPerformed_2017\Lab Data Oct-2017_Sample-collected-orders_Tests.xlsx";
            List<string> sheetNames = new List<string>()
            {
                "Lab_Order_Collected _Samples",
                "Lab Test at PHCC_Oct 2017"
            };

            IExcelOperations excelOps = new OpenXMLDOMLib();
            var sheets = excelOps.GetSheetNames(path);
            for (int i = 0; i < sheetNames.Count; i++)
            {
                Assert.Contains(sheetNames[i], sheets);
            }
        }
    }
}
