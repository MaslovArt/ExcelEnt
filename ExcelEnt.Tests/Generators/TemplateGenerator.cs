using ExcelEnt.Tests.Models;
using ExcelEnt.Write;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace ExcelEnt.Tests.Generators
{
    public class TemplateGenerator
    {
        private string redBGStyle = "table-red-bg";
        private string blueBGStyle = "table-blue-bg";
        private string tableStyle = "table";

        public XSSFWorkbook Generate(string templatePath, TestItem[] items, int offset = 0)
        {
            var generator = new XLSXWriter<TestItem>();
            return generator
                .AddRule(e => e.Int, offset + 0)
                .AddRule(e => e.NullInt, offset + 1)
                .AddRule(e => e.Double, offset + 2)
                .AddRule(e => e.NullDouble, offset + 3)
                .AddRule(e => e.Date, offset + 4)
                .AddRule(e => e.NullDate, offset + 5)
                .AddRule(e => e.Bool, offset + 6)
                .AddRule(e => e.NullBool, offset + 7)
                .AddRule(e => e.String, offset + 8)
                .AddRule(e => e.Enum, offset + 9)
                .UseTemplating(t => t
                    .FromTemplate(templatePath, 3, true)
                    .ReplaceShortCode("header", "inserted-header-val-1")
                    .ReplaceShortCode("footer", "inserted-footer-val-1"))
                .UseStyling(s => s
                    .AddStyle(CreateRedStyle, redBGStyle)
                    .AddStyle(CreateTableStyle, tableStyle)
                    .AddStyle(CreateBlueStyle, blueBGStyle)
                    .AddCellsStyle(blueBGStyle, 0)
                    .AddCellsStyle(tableStyle)
                    .AddConditionRowStyle((e, i) => e.NullInt == null ? redBGStyle : null))
                .Generate(items);
        }

        private void CreateRedStyle(ICellStyle style)
        {
            style.BorderBottom = BorderStyle.Thin;
            style.BorderTop = BorderStyle.Thin;
            style.BorderLeft = BorderStyle.Thin;
            style.BorderRight = BorderStyle.Thin;
            style.FillForegroundColor = IndexedColors.Red.Index;
            style.FillPattern = FillPattern.SolidForeground;
        }

        private void CreateTableStyle(ICellStyle style)
        {
            style.BorderBottom = BorderStyle.Thin;
            style.BorderTop = BorderStyle.Thin;
            style.BorderLeft = BorderStyle.Thin;
            style.BorderRight = BorderStyle.Thin;
        }

        private void CreateBlueStyle(ICellStyle style)
        {
            style.BorderBottom = BorderStyle.Thin;
            style.BorderTop = BorderStyle.Thin;
            style.BorderLeft = BorderStyle.Thin;
            style.BorderRight = BorderStyle.Thin;
            style.FillForegroundColor = IndexedColors.LightBlue.Index;
            style.FillPattern = FillPattern.SolidForeground;
        }
    }
}
