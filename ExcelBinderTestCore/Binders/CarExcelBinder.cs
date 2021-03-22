using ExcelBinderTestCore.Models;
using ExcelEnt.Bind;

namespace ExcelBinderTestCore.Readers
{
    public class CarExcelBinder
    {
        public Car[] Bind(string filePath, int pageInd)
        {
            var excelCars = new XLSXBinder<Car>()
                .AddRule(0, m => m.Brand, BindMappers.String)
                .AddRule(1, m => m.Model, BindMappers.String)
                .AddRule(2, m => m.Year, BindMappers.NullInt)
                .AddRule(3, m => m.HP, BindMappers.NullInt)
                .AddRule(4, m => m.Crashed, (cell) => BindMappers.StringBool(cell, "Yes"))
                .AddRule(5, m => m.Class, BindMappers.Enum<CarClass>)
                .Skip(1)
                .Bind(filePath, pageInd);

            return excelCars;
        }
    }
}
