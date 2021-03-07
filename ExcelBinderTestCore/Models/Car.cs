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
        [WriteCol(2)]
        [ReadIntCol(2)]
        public int? Year { get; set; }

        [Description("HP")]
        [WriteCol(3)]
        [ReadIntCol(3)]
        public int? HP { get; set; }

        [ReadBoolCol(4, TrueValue = "Yes")]
        public bool Crashed { get; set; }

        [Description("Crashed")]
        [WriteCol(4)]
        public string CrashedStr => Crashed ? "Yes" : "No";

        [Description("Class")]
        [WriteCol(5)]
        [ReadEnumCol(5, typeof(CarClass))]
        public CarClass Class { get; set; }

        public override string ToString()
        {
            return $"Car [{Brand} {Model} {Year}], HP={HP}, Crashed={Crashed}, Class={Class}";
        }
    }
}
