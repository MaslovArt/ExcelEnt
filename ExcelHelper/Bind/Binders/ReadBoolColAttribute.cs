using NPOI.SS.UserModel;

namespace ExcelHelper.Bind.Binders
{
    public class ReadBoolColAttribute : BaseColAttribute
    {
        public ReadBoolColAttribute(int columnIndex)
            : base(columnIndex) { }

        public string TrueValue { get; set; }

        protected override object ParseValue(ICell value)
        {
            if (string.IsNullOrEmpty(TrueValue))
            {
                return value.BooleanCellValue;
            }

            return value.StringCellValue.ToUpper() == TrueValue.ToUpper();
        }
    }
}
