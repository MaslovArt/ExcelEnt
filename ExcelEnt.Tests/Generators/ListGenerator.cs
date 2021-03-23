using ExcelEnt.Tests.Models;
using ExcelEnt.Write;
using NPOI.XSSF.UserModel;

namespace ExcelEnt.Tests.Generators
{
    public class ListGenerator
    {
        public XSSFWorkbook Generate(TestItem[] items, int offset = 0)
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
                .Generate(items);
        }
    }
}
