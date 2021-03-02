using NPOI.SS.UserModel;

namespace ExcelHelper.Bind.Binders
{
    public class DoubleColAttribute : BaseColAttribute
    {
        public DoubleColAttribute(int columnIndex)
            : base(columnIndex) { }

        protected override object ParseValue(ICell value) => value.NumericCellValue;
    }
}
