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
            var cars = new Car[]
            {
                new Car()
                {
                    Brand = "BWM",
                    Model = "3",
                    Crashed = false,
                    HP = 122,
                    Year = 2016,
                    Class = CarClass.C
                },
                new Car()
                {
                    Brand = "BWM",
                    Model = "5",
                    Crashed = false,
                    HP = 250,
                    Year = 2016,
                    Class = CarClass.S
                },
                new Car()
                {
                    Brand = "BWM",
                    Model = "M4",
                    Crashed = false,
                    HP = 450,
                    Year = 2016,
                    Class = CarClass.E
                }
            };

            var excelGenerator = new XLSXWriter();
            excelGenerator.Generate(carsReportTemplate, "Лист1", carsReportFile, 4, cars, true);

            Console.ReadKey();
        }

        static void Print(IEnumerable<Car> cars)
        {
            foreach (var car in cars)
                Console.WriteLine(car);
        }
    }
}
