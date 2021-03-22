using System;

namespace ExcelEnt.Tests.Models
{
    public class TestItem
    {
        public int Int { get; set; }
        public int? NullInt { get; set; }
        public double Double { get; set; }
        public double? NullDouble { get; set; }
        public DateTime Date { get; set; }
        public DateTime? NullDate { get; set; }
        public bool Bool { get; set; }
        public bool? NullBool { get; set; }
        public string String { get; set; }
        public TestEnum Enum { get; set; }
        public object Obj { get; set; }

        public override bool Equals(object obj)
        {
            if (obj is TestItem testItem)
            {
                return
                    Int == testItem.Int &&
                    NullInt == testItem.NullInt &&
                    Double == testItem.Double &&
                    NullDouble == testItem.NullDouble &&
                    Bool == testItem.Bool &&
                    NullBool == testItem.NullBool &&
                    Date == testItem.Date &&
                    NullDate == testItem.NullDate &&
                    String == testItem.String &&
                    Enum == testItem.Enum;
            }

            return false;
        }

        public override int GetHashCode()
        {
            return base.GetHashCode();
        }
    }
}
