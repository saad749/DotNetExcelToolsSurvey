using DotNetExcelToolsSurvey.Library;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DotnetExcelToolsSurvey.Console
{
    class Program
    {
        static void Main(string[] args)
        {
            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();
            IExcelOperations excelOps = new ExcelDataReaderLib();
            //var values = excelOps.GetSheetNames(
            //    @"C:\Users\PHCCUser\Desktop\PMS Data\2017\DW_2017\LAB_TestsPerformed_2017\Lab Data Oct-2017_Sample-collected-orders_Tests.xlsx");
            var values = excelOps.GetFirstRow(
                @"C:\Users\PHCCUser\Desktop\PMS Data\2017\DW_2017\LAB_TestsPerformed_2017\Lab Data Oct-2017_Sample-collected-orders_Tests.xlsx");

            stopWatch.Stop();
            PrintHeaders(values);


            System.Console.WriteLine(stopWatch.ElapsedMilliseconds);
            System.Console.ReadLine();
        }

        static void PrintSheets(List<string> values)
        {
            foreach (var value in values)
            {
                System.Console.WriteLine(value);
            }
        }
        static void PrintHeaders(List<List<string>> values)
        {
            foreach (var value in values)
            {
                foreach (var item in value)
                {
                    System.Console.WriteLine(item);
                }
                System.Console.WriteLine();

            }
        }


    }
}
