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
            var carsBinder = new CarExcelBinder();
            var excelCars = carsBinder.Bind(@"Documents\Cars.xlsx", 0);

            Print(excelCars);

            new CarsReportGenerator().Generate("cars_report.xlsx", excelCars);
        }

        static void Print(IEnumerable<Car> cars)
        {
            foreach (var car in cars)
                Console.WriteLine(car);
        }
    }
}
