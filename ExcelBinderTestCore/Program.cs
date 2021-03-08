using ExcelBinderTestCore.Models;
using ExcelHelper.Bind;
using ExcelHelper.Write;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;

namespace ExcelBinderTestCore
{
    class Program
    {
        static void Main(string[] args)
        {
            var carsFile = @"Documents\Cars.xlsx";

            var excelCars = new XLSXBinder<Car>()
                .AddRule(0, m => m.Brand, BindMappers.String)
                .AddRule(1, m => m.Model, BindMappers.String)
                .AddRule(2, m => m.Year, BindMappers.NullInt)
                .AddRule(3, m => m.HP, BindMappers.NullInt)
                .AddRule(4, m => m.Crashed, (cell) => BindMappers.StringBool(cell, "Yes"))
                .AddRule(5, m => m.Class, BindMappers.Enum<CarClass>)
                .StartFrom(1)
                .Bind(carsFile, 0);

            Print(excelCars);


            var carsReportTemplate = @"Documents\CarsReport.xlsx";
            var carsReportFile = "result.xlsx";

            //new XLSXWriter<Car>()
            //    .UseModelDescription()
            //    .Generate(carsReportFile, excelCars);

            //new XLSXWriter<Car>()
            //    .UseTemplate(carsReportTemplate, 0, 4, true)
            //    .Generate(carsReportFile, excelCars);
        }

        static void Print(IEnumerable<Car> cars)
        {
            foreach (var car in cars)
                Console.WriteLine(car);
        }
    }
}
