using ExcelBinderTestCore.Models;
using ExcelHelper.Bind;
using System;
using System.Collections.Generic;
using System.IO;

namespace ExcelBinderTestCore
{
    class Program
    {
        static void Main(string[] args)
        {
            var file = new FileInfo(@"Documents\Cars.xlsx");
            var excelBinder = new XLSXBinder();

            var excelCars = excelBinder.Bind<Car>(file, "Лист1", 1);
            Print(excelCars);

            Console.ReadKey();
        }

        static void Print(IEnumerable<Car> cars)
        {
            foreach (var car in cars)
                Console.WriteLine(car);
        }
    }
}
