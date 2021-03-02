using NPOI.SS.UserModel;

namespace ExcelHelper.Binders
{
    public class BoolColAttribute : BaseColAttribute
    {
        public BoolColAttribute(int columnIndex, string trueValue)
            : base(columnIndex)
        {
            TrueValue = trueValue;
        }

        public string TrueValue { get; private set; }

        protected override object ParseValue(ICell value) =>
            value.StringCellValue.ToUpper() == TrueValue.ToUpper();
    }
}
