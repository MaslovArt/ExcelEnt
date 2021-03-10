using ExcelBinderTestCore.Models;
using ExcelEnt.Write;
using NPOI.SS.UserModel;
using System.Collections.Generic;
using System.Linq;

namespace ExcelBinderTestCore.Generarors
{
    public class CarsReportGenerator
    {
        private string RedBGStyle = "table-red-bg";
        private string TableStyle = "table";
        private string templatePath = @"Documents\CarsReport.xlsx";

        public void Generate(string newFile, IEnumerable<Car> cars)
        {
            var generator = new XLSXWriter<Car>();
            var carsArray = cars.ToArray();

            generator
                .FromTemplate(templatePath, 0, 4, true)
                .ReplaceShortCode("CarCount", carsArray.Length.ToString())
                .ReplaceShortCode("Name", "Putin V.V.")
                .AddRule(c => c.Brand, 0)
                .AddRule(c => c.Model, 1)
                .AddRule(c => c.Year, 2)
                .AddRule(c => c.HP, 3)
                .AddRule(c => c.Crashed, 4)
                .AddRule(c => c.Class, 5)
                .AddStyle(CreateRedStyle, RedBGStyle)
                .AddStyle(CreateTableStyle, TableStyle)
                .AddDefaultRowsStyle(TableStyle)
                .AddConditionRowStyle(car => car.Crashed ? RedBGStyle : null)
                .Generate(newFile, carsArray);
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
    }
}
