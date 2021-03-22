using ExcelEnt.Tests.Models;
using System;

namespace ExcelEnt.Tests.ExpectedData
{
    public static class Expected
    {
        public static TestItem[] Data => new TestItem[]
        {
            new TestItem()
            {
                Int = 1,
                NullInt = 1,
                Double = 2.5,
                NullDouble = 2.5,
                Date = new DateTime(2020, 1, 1, 1, 10, 10),
                NullDate = new DateTime(2020, 1, 1, 1, 10, 10),
                Bool = true,
                NullBool = true,
                String = "Str value",
                Enum = TestEnum.Value1,
                Obj = new object()
            },
            new TestItem()
            {
                Int = 1,
                NullInt = null,
                Double = 2.5,
                NullDouble = null,
                Date = new DateTime(2020, 1, 1, 1, 10, 10),
                NullDate = null,
                Bool = false,
                NullBool = false,
                String = "Str value nf",
                Enum = TestEnum.Value3,
                Obj = new object()
            },
            new TestItem()
            {
                Int = 1,
                NullInt = 1,
                Double = 2.5,
                NullDouble = 2.5,
                Date = new DateTime(2020, 1, 1, 1, 10, 10),
                NullDate = new DateTime(2020, 1, 1, 1, 10, 10),
                Bool = true,
                NullBool = true,
                String = "Str value",
                Enum = TestEnum.Value1,
                Obj = new object()
            },
            new TestItem()
            {
                Int = 1,
                NullInt = null,
                Double = 2.5,
                NullDouble = null,
                Date = new DateTime(2020, 1, 1, 1, 10, 10),
                NullDate = null,
                Bool = false,
                NullBool = false,
                String = "Str value nf",
                Enum = TestEnum.Value3,
                Obj = new object()
            },
        };
    }
}
