using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace DotNetExcelToolsSurvey.Tests
{
    public class ReadSheetTests
    {
        [Fact]
        public void PassingTest()
        {
            Assert.Equal(4, 2+2);
        }

        [Theory]
        [InlineData("", )]
        [InlineData(5)]
        [InlineData(6)]
        public void IsReadingSheetNames(string path, List<string> sheetNames)
        {
            Assert.Equal(4, 2 + 2);
        }


    }
}
