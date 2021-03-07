using NPOI.SS.UserModel;

namespace ExcelHelper.Bind.Binders
{
    public class ReadDoubleColAttribute : BaseColAttribute
    {
        public ReadDoubleColAttribute(int columnIndex)
            : base(columnIndex) { }

        protected override object ParseValue(ICell value) => value.NumericCellValue;
    }
}
