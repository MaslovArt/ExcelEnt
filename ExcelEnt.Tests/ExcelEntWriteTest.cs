using ExcelEnt.Tests.ExpectedData;
using ExcelEnt.Tests.Generators;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace ExcelEnt.Tests
{
    [TestClass]
    public class ExcelEntWriteTest
    {
        [TestMethod]
        public void ListWriteTest()
        {
            var expectedWb = new XSSFWorkbook(Expected.ListPath);

            var listGenerator = new ListGenerator();
            var generatedWb = listGenerator.Generate(Expected.Data);

            AssertWorkbookEquals(expectedWb, generatedWb);
        }

        [TestMethod]
        public void TitleListWriteTest()
        {
            var expectedWb = new XSSFWorkbook(Expected.TitleListPath);

            var listGenerator = new TitleListGenerator();
            var generatedWb = listGenerator.Generate(Expected.Data);

            AssertWorkbookEquals(expectedWb, generatedWb);
        }

        [TestMethod]
        public void OffsetTitleListWriteTest()
        {
            var expectedWb = new XSSFWorkbook(Expected.TitleListOffset3Path);

            var listGenerator = new TitleListGenerator();
            var generatedWb = listGenerator.Generate(Expected.Data, 3);

            AssertWorkbookEquals(expectedWb, generatedWb);
        }

        [TestMethod]
        public void TemplateWriteTest()
        {
            var expectedWb = new XSSFWorkbook(Expected.TemplateListPath);

            var listGenerator = new TemplateGenerator();
            var generatedWb = listGenerator.Generate(Expected.GenTemplatePath, Expected.Data);

            AssertWorkbookEquals(expectedWb, generatedWb);
        }


        private void AssertWorkbookEquals(XSSFWorkbook wb1, XSSFWorkbook wb2)
        {
            var wb1Sheet = wb1.GetSheetAt(0);
            var wb2Sheet = wb2.GetSheetAt(0);

            Assert.IsTrue(wb1Sheet.LastRowNum == wb2Sheet.LastRowNum);

            foreach (IRow row in wb1Sheet)
            {
                foreach (ICell cell in row)
                {
                    var wb2SameCell = wb2Sheet
                        .GetRow(cell.RowIndex)
                        .GetCell(cell.ColumnIndex);

                    var styleEq = cell.CellStyle.Equals(wb2SameCell.CellStyle);
                    var typeEq = cell.CellType == wb2SameCell.CellType;
                    var valueEq = cell.ToString() == wb2SameCell.ToString();

                    var messageErr =
                        $"{(!styleEq ? "style " : null)}" +
                        $"{(!typeEq ? "type " : null)}" +
                        $"{(!valueEq ? "value" : null)}";
                    messageErr = messageErr.Trim().Replace(" ", ", ");

                    var equal = styleEq && typeEq && valueEq;

                    Assert.IsTrue(equal, $"Cells [{cell.RowIndex},{cell.ColumnIndex}] {messageErr} not equal.");
                }
            }
        }
    }
}
