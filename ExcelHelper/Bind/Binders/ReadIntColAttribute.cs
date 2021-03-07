using NPOI.SS.UserModel;

namespace ExcelHelper.Bind.Binders
{
    public class ReadIntColAttribute : BaseColAttribute
    {
        public ReadIntColAttribute(int columnIndex)
            : base(columnIndex) { }

        protected override object ParseValue(ICell value) => (int)value.NumericCellValue;
    }
}
