using NPOI.SS.UserModel;

namespace ExcelHelper.Bind.Binders
{
    public class GetBoolColAttribute : BaseColAttribute
    {
        public GetBoolColAttribute(int columnIndex, string trueValue)
            : base(columnIndex)
        {
            TrueValue = trueValue;
        }

        public string TrueValue { get; set; }

        protected override object ParseValue(ICell value) =>
            value.StringCellValue.ToUpper() == TrueValue.ToUpper();
    }
}
