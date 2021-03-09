using ExcelBinderTestCore.Models;
using ExcelHelper.Write;
using NPOI.SS.UserModel;
using System.Collections.Generic;
using System.Linq;

namespace ExcelBinderTestCore.Generarors
{
    public class CarsListGenerator
    {
        const string RED_BG_STYLE = "red-bg";
        const string GRAY_BG_STYLE = "gray-bg";

        public void Generate(string newFile, IEnumerable<Car> cars)
        {
            var generator = new XLSXWriter<Car>();
            generator
                .AddRule(c => c.Brand, 0, GRAY_BG_STYLE)
                .AddRule(c => c.Model, 1)
                .AddRule(c => c.Year, 2)
                .AddRule(c => c.HP, 3)
                .AddRule(c => c.Crashed, 4)
                .AddRule(c => c.Class, 6)
                .AddStyle(CreateRedBGStyle, RED_BG_STYLE)
                .AddStyle(CreateGrayStyle, GRAY_BG_STYLE)
                .AddConditionRowStyle(car => car.Crashed ? RED_BG_STYLE : null)
                .Generate(newFile, cars.ToArray());
        }

        private void CreateRedBGStyle(ICellStyle style)
        {
            style.FillForegroundColor = IndexedColors.Red.Index;
            style.FillPattern = FillPattern.SolidForeground;
        }

        private void CreateGrayStyle(ICellStyle style)
        {
            style.FillForegroundColor = IndexedColors.Grey25Percent.Index;
            style.FillPattern = FillPattern.SolidForeground;
        }
    }
}
