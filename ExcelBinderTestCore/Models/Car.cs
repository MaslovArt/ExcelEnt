using ExcelHelper.Bind.Binders;

namespace ExcelBinderTestCore.Models
{
    public class Car
    {
        [StringCol(0)]
        public string Brand { get; set; }

        [StringCol(1)]
        public string Model { get; set; }

        [IntCol(2)]
        public int Year { get; set; }

        [IntCol(3)]
        public int HP { get; set; }

        [BoolCol(4, "Yes")]
        public bool Crashed { get; set; }

        [EnumCol(5, typeof(CarClass))]
        public CarClass Class { get; set; }

        public override string ToString()
        {
            return $"Car [{Brand} {Model} {Year}], HP={HP}, Crashed={Crashed}, Class={Class}";
        }
    }
}
