namespace ExcelBinderTestCore.Models
{
    public class Car
    {
        public string Brand { get; set; }

        public string Model { get; set; }

        public int? Year { get; set; }

        public int? HP { get; set; }

        public bool Crashed { get; set; }

        public CarClass Class { get; set; }

        public override string ToString()
        {
            return $"Car [{Brand} {Model} {Year}], HP={HP}, Crashed={Crashed}, Class={Class}";
        }
    }
}
