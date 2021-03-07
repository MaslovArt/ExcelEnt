using ExcelBinderTestCore.Models;
using ExcelHelper.Bind;
using ExcelHelper.Write;
using System;
using System.Collections.Generic;

namespace ExcelBinderTestCore
{
    class Program
    {
        static void Main(string[] args)
        {
            var carsFile = @"Documents\Cars.xlsx";

            var excelBinder = new XLSXBinder();
            var excelCars = excelBinder.Bind<Car>(carsFile, "Лист1", 1);
            Print(excelCars);


            var carsReportTemplate = @"Documents\CarsReport.xlsx";
            var carsReportFile = "result.xlsx";

            //new XLSXWriter<Car>()
            //    .UseModelDescription()
            //    .Generate(carsReportFile, excelCars);

            new XLSXWriter<Car>()
                .UseTemplate(carsReportTemplate, 0, 4, true)
                .Generate(carsReportFile, excelCars);

            Console.ReadKey();
        }

        static void Print(IEnumerable<Car> cars)
        {
            foreach (var car in cars)
                Console.WriteLine(car);
        }
    }
}
