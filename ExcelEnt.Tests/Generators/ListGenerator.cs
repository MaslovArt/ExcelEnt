using ExcelEnt.Tests.Models;
using ExcelEnt.Write;

namespace ExcelEnt.Tests.Generators
{
    public class ListGenerator
    {
        public void Generate(string filePath, TestItem[] items)
        {
            var generator = new XLSXWriter<TestItem>();
            generator
                .AddRule(e => e.Int, 0)
                .AddRule(e => e.NullInt, 1)
                .AddRule(e => e.Double, 2)
                .AddRule(e => e.NullDouble, 3)
                .AddRule(e => e.Date, 4)
                .AddRule(e => e.NullDate, 5)
                .AddRule(e => e.Bool, 6)
                .AddRule(e => e.NullBool, 7)
                .AddRule(e => e.String, 8)
                .AddRule(e => e.Enum, 9)
                .Generate(filePath, items);
        }
    }
}
