using ExcelBinderTestCore.Models;
using ExcelEnt.Write;
using NPOI.SS.UserModel;
using System.Collections.Generic;
using System.Linq;

namespace ExcelBinderTestCore.Generarors
{
    public class CarsReportGenerator
    {
        private string redBGStyle = "table-red-bg";
        private string blueBGStyle = "table-blue-bg";
        private string tableStyle = "table";
        private string templatePath = @"Documents\CarsReport.xlsx";

        public void Generate(string newFile, IEnumerable<Car> cars)
        {
            var generator = new XLSXWriter<Car>();
            var carsArray = cars.ToArray();

            generator
                .AddRule(c => c.Brand, 0)
                .AddRule(c => c.Model, 1)
                .AddRule(c => c.Year, 2)
                .AddRule(c => c.HP, 3)
                .AddRule(c => c.Crashed, 4)
                .AddRule(c => c.Class, 5)
                .UseTemplating(t => t
                    .FromTemplate(templatePath, 4, true)
                    .ReplaceShortCode("CarCount", carsArray.Length.ToString())
                    .ReplaceShortCode("Name", "Putin V.V."))
                .UseStyling(s => s
                    .AddStyle(CreateRedStyle, redBGStyle)
                    .AddStyle(CreateTableStyle, tableStyle)
                    .AddStyle(CreateBlueStyle, blueBGStyle)
                    .AddCellStyle(0, blueBGStyle)
                    .AddCellsDefaultStyle(tableStyle)
                    .AddConditionRowStyle((car, i) => car.Crashed ? redBGStyle : null))
                .Generate(carsArray)
                .SaveTo(newFile);
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
