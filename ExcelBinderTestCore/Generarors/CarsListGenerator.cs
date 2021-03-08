using ExcelBinderTestCore.Models;
using ExcelHelper.Write;
using System.Collections.Generic;
using System.Linq;

namespace ExcelBinderTestCore.Generarors
{
    public class CarsListGenerator
    {
        public void Generate(string newFile, IEnumerable<Car> cars)
        {
            var generator = new XLSXWriter<Car>();
            generator
                .FromEmptyWithHeaders(new string[]
                {
                    "1", "2", "3", "4", "5", "6"
                })
                .AddRule(c => c.Brand, 0)
                .AddRule(c => c.Model, 1)
                .AddRule(c => c.Year, 2)
                .AddRule(c => c.HP, 3)
                .AddRule(c => c.Crashed, 4)
                .AddRule(c => c.Class, 5)
                .Generate(newFile, cars.ToArray());
        }
    }
}
