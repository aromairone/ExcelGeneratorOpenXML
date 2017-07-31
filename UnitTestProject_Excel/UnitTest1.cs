using ExcelGeneratorOpenXML;
using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelGeneratorOpenXML.Tests
{
    [TestClass()]
    public class UnitTest1
    {
        [TestMethod()]
        public void CreateSpreadSheetTest()
        {
            ExcelService service = new ExcelService();
            bool isActual = service.CreateSpreadSheet();
            bool isExpected = true;
            Assert.AreEqual(isExpected,isActual);
        }
    }
}
