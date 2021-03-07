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


            //var carsReportTemplate = @"Documents\CarsReport.xlsx";
            //var carsReportFile = "result.xlsx";

            //var excelGenerator = new XLSXWriter();
            //excelGenerator.Generate(carsReportTemplate, "Лист1", carsReportFile, 4, excelCars, true);

            Console.ReadKey();
        }

        static void Print(IEnumerable<Car> cars)
        {
            foreach (var car in cars)
                Console.WriteLine(car);
        }
    }
}
