using ExcelEnt.Tests.ExpectedData;
using ExcelEnt.Tests.Generators;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelEnt.Tests
{
    [TestClass]
    public class ExcelEntWriteTest
    {
        [TestMethod]
        public void ListWriteTest()
        {
            var listResultPath = @"listResult.xlsx";
            var listGenerator = new ListGenerator();
            listGenerator.Generate(listResultPath, Expected.Data);
        }
    }
}
