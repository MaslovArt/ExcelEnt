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
            var carsReportFile = "result.xlsx";

            var carsBinder = new CarExcelBinder();
            var excelCars = carsBinder.Bind(carsFile, 0);
            Print(excelCars);

            var carsReportGenerator = new CarsReportGenerator();
            carsReportGenerator.Generate(carsReportFile, excelCars);
        }

        static void Print(IEnumerable<Car> cars)
        {
            foreach (var car in cars)
                Console.WriteLine(car);
        }
    }
}
