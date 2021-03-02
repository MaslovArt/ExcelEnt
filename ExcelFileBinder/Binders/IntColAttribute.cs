using NPOI.SS.UserModel;

namespace ExcelHelper.Binders
{
    public class IntColAttribute : BaseColAttribute
    {
        public IntColAttribute(int columnIndex)
            : base(columnIndex) { }

        protected override object ParseValue(ICell value) => (int)value.NumericCellValue;
    }
}
