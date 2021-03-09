using ExcelBinderTestCore.Generarors;
using ExcelBinderTestCore.Models;
using ExcelBinderTestCore.Readers;
using System;
using System.Collections.Generic;

namespace ExcelBinderTestCore
{
    class Program
    {
        static void Main(string[] args)
        {
            var carsFile = @"Documents\Cars.xlsx";
            var carsReportFile = "reportCarsResult.xlsx";
            var carsListFile = "listCarsResult.xlsx";

            var carsBinder = new CarExcelBinder();
            var excelCars = carsBinder.Bind(carsFile, 0);
            Print(excelCars);

            var carsReportGenerator = new CarsReportGenerator();
            carsReportGenerator.Generate(carsReportFile, excelCars);

            //var carsListGenerator = new CarsListGenerator();
            //carsListGenerator.Generate(carsListFile, excelCars);
        }

        static void Print(IEnumerable<Car> cars)
        {
            foreach (var car in cars)
                Console.WriteLine(car);
        }
    }
}
