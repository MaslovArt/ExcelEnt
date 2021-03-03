using ExcelHelper.Bind.Binders;
using ExcelHelper.Write;

namespace ExcelBinderTestCore.Models
{
    public class Car
    {
        [WriteCol(0)]
        [GetStringCol(0)]
        public string Brand { get; set; }

        [WriteCol(1)]
        [GetStringCol(1)]
        public string Model { get; set; }

        [WriteCol(2)]
        [GetIntCol(2)]
        public int Year { get; set; }

        [WriteCol(3)]
        [GetIntCol(3)]
        public int HP { get; set; }

        [WriteCol(4)]
        [GetBoolCol(4, "Yes")]
        public bool Crashed { get; set; }

        [WriteCol(5)]
        [GetEnumCol(5, typeof(CarClass))]
        public CarClass Class { get; set; }

        public override string ToString()
        {
            return $"Car [{Brand} {Model} {Year}], HP={HP}, Crashed={Crashed}, Class={Class}";
        }
    }
}
