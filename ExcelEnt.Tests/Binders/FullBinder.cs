using ExcelEnt.Bind;
using ExcelEnt.Tests.Models;

namespace ExcelEnt.Tests.Binders
{
    public class FullBinder
    {
        public TestItem[] Bind(string filePath)
        {
            var excelCars = new XLSXBinder<TestItem>()
                .AddRule(0, e => e.Int, BindMappers.Int)
                .AddRule(1, e => e.NullInt, BindMappers.NullInt)
                .AddRule(2, e => e.Double, BindMappers.Double)
                .AddRule(3, e => e.NullDouble, BindMappers.NullDouble)
                .AddRule(4, e => e.Date, BindMappers.Date)
                .AddRule(5, e => e.NullDate, BindMappers.NullDate)
                .AddRule(6, e => e.Bool, BindMappers.Bool)
                .AddRule(7, e => e.NullBool, BindMappers.NullBool)
                .AddRule(8, e => e.String, BindMappers.String)
                .AddRule(9, e => e.Enum, BindMappers.Enum<TestEnum>)
                .Bind(filePath, 0);

            return excelCars;
        }
    }
}
