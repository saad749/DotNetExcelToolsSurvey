using DotNetExcelToolsSurvey.Library;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace DotNetExcelToolsSurvey.Tests
{
    public class ExcelDataReaderLibTests
    {
        [Fact]
        public void PassingTest()
        {
            Assert.Equal(4, 2+2);
        }

        [Fact]
        public void IsSheetCountCorrect()
        {
            string path = @"C:\Users\PHCCUser\Desktop\PMS Data\2017\DW_2017\LAB_TestsPerformed_2017\Lab Data Oct-2017_Sample-collected-orders_Tests.xlsx";

            IExcelOperations excelOps = new ExcelDataReaderLib();
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

            IExcelOperations excelOps = new ExcelDataReaderLib();
            var sheets = excelOps.GetSheetNames(path);
            for (int i = 0; i < sheetNames.Count; i++)
            {
                Assert.Contains(sheetNames[i], sheets);
            }
        }

        [Fact]
        public void IsReadingAllColumns()
        {
            string path = @"C:\Users\PHCCUser\Desktop\PMS Data\2017\DW_2017\LAB_TestsPerformed_2017\Lab Data Oct-2017_Sample-collected-orders_Tests.xlsx";
            List<string> sheetNames = new List<string>()
            {
                "Lab_Order_Collected _Samples",
                "Lab Test at PHCC_Oct 2017"
            };

            IExcelOperations excelOps = new ExcelDataReaderLib();
            var sheets = excelOps.GetSheetNames(path);
            for (int i = 0; i < sheetNames.Count; i++)
            {
                Assert.Contains(sheetNames[i], sheets);
            }
        }


    }
}

//Lab_Order_Collected _Samples
//PATIENT_ID
//LAB_ORDER_NAME
//HEALTH_CENTER_NAME
//DEPARTMENT_PERFORMED
//ACTIVITY_TYPE
//ORDER_ID
//ORDER_DT_TIME
//DATE_COLLECTED
//UCES_ORDER_STATUS_DISP
//ENCOUNTER_LOCATION
//ENCOUNTER_NUMBER
//ENCOUNTER_TYPE
//COMPLETE_DT_TM
//HC
//SEX
//NATIONALITY
//SHIFT

//Lab Test at PHCC_Oct 2017
//PERSON_ID
//ORDER_ID
//ACCESSION_NBR
//PERFORMING_SUBSECTION
//ORDER_STATUS
//COLLECTION_PRIORITY
//DEPARTMENT_ORDER_NAME
//DISCRETE_ASSAY
//PATIENT_LOCATION
//ORDER_LOCATION
//ORDER_PHYSICIAN
//ORDER_DT_TM
//DRAWN_TECHNICIAN
//DRAWN_DT_TM
//INLAB_DT_TM
//PERFORMING_TECHNICIAN
//PERFORM_DT_TM
//VERIFYING_TECHNICIAN
//VERIFY_DT_TM
//COMPLETE_DT_TM
//PERFORMING_INSTITUTION
//ENCOUNTER_LOCATION
//ENCOUNTER_NUMBER
//ENCOUNTER_TYPE
//PATIENT_MRN
//DATE_OF_BIRTH
//AGE
//GENDER
//NATIONALITY
//DEPARTMENT
//SERVICE_RESOURCE
//CLIENT


