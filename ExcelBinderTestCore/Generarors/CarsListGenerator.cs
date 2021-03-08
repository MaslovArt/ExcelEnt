using ExcelBinderTestCore.Models;
using ExcelHelper.Write;
using NPOI.SS.UserModel;
using System.Collections.Generic;
using System.Linq;

namespace ExcelBinderTestCore.Generarors
{
    public class CarsListGenerator
    {
        const string GRAY_BG_STYLE = "gray-bg";

        public void Generate(string newFile, IEnumerable<Car> cars)
        {
            var generator = new XLSXWriter<Car>();
            generator
                .FromEmptyWithHeaders(new string[]
                {
                    "1", "2", "3", "4", "5", "6"
                })
                .AddRule(c => c.Brand, 0, GRAY_BG_STYLE)
                .AddRule(c => c.Model, 1)
                .AddRule(c => c.Year, 2)
                .AddRule(c => c.HP, 3)
                .AddRule(c => c.Crashed, 4)
                .AddRule(c => c.Class, 5)
                .AddStyle(CreateGrayStyle, GRAY_BG_STYLE)
                .Generate(newFile, cars.ToArray());
        }

        private void CreateGrayStyle(ICellStyle style)
        {
            style.FillForegroundColor = IndexedColors.Grey25Percent.Index;
            style.FillPattern = FillPattern.SolidForeground;
        }
    }
}
