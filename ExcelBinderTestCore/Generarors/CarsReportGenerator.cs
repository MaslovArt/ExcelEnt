using ExcelBinderTestCore.Models;
using ExcelHelper.Write;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelBinderTestCore.Generarors
{
    public class CarsReportGenerator
    {
        private string templatePath = @"Documents\CarsReport.xlsx";
        private int templatePage = 0;
        private int insertInd = 4;
        private bool moveFooter = true;

        public void Generate(string newFile, IEnumerable<Car> cars)
        {
            var generator = new XLSXWriter<Car>();
            generator
                .FromTemplate(templatePath, templatePage, insertInd, moveFooter)
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
