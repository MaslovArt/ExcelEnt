using ExcelBinderTestCore.Models;
using ExcelHelper.Write;
using NPOI.SS.UserModel;
using System.Collections.Generic;
using System.Linq;

namespace ExcelBinderTestCore.Generarors
{
    public class CarsReportGenerator
    {
        private string RedBGStyle = "red-bg";
        private string templatePath = @"Documents\CarsReport.xlsx";
        private int templatePage = 0;
        private int insertInd = 4;
        private bool moveFooter = true;

        public void Generate(string newFile, IEnumerable<Car> cars)
        {
            var generator = new XLSXWriter<Car>();
            var carsArray = cars.ToArray();

            generator
                .FromTemplate(templatePath, templatePage, insertInd, moveFooter)
                .ReplaceShortCode("CarCount", carsArray.Length.ToString())
                .ReplaceShortCode("Name", "Putin V.V.")
                .AddRule(c => c.Brand, 0)
                .AddRule(c => c.Model, 1)
                .AddRule(c => c.Year, 2)
                .AddRule(c => c.HP, 3)
                .AddRule(c => c.Crashed, 4)
                .AddRule(c => c.Class, 5)
                .AddStyle(CreateRedStyle, RedBGStyle)
                .AddConditionRowStyle(car => car.Crashed ? RedBGStyle : null)
                .Generate(newFile, carsArray);
        }

        private void CreateRedStyle(ICellStyle style)
        {
            style.FillForegroundColor = IndexedColors.Red.Index;
            style.FillPattern = FillPattern.SolidForeground;
        }
    }
}
