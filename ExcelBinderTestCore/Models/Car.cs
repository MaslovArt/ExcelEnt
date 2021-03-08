using ExcelHelper.Bind.Binders;
using ExcelHelper.Write;
using System.ComponentModel;

namespace ExcelBinderTestCore.Models
{
    public class Car
    {
        [Description("Brand")]
        [WriteCol(0)]
        [ReadStringCol(0)]
        public string Brand { get; set; }

        [Description("Model")]
        [WriteCol(1)]
        [ReadStringCol(1)]
        public string Model { get; set; }

        [Description("Year")]
        public int? Year { get; set; }

        [Description("HP")]
        public int? HP { get; set; }

        public bool Crashed { get; set; }

        [Description("Crashed")]
        public string CrashedStr => Crashed ? "Yes" : "No";

        [Description("Class")]
        public CarClass Class { get; set; }

        public override string ToString()
        {
            return $"Car [{Brand} {Model} {Year}], HP={HP}, Crashed={Crashed}, Class={Class}";
        }
    }
}
