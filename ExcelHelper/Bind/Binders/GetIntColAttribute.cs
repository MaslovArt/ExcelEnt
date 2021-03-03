using NPOI.SS.UserModel;

namespace ExcelHelper.Bind.Binders
{
    public class GetIntColAttribute : BaseColAttribute
    {
        public GetIntColAttribute(int columnIndex)
            : base(columnIndex) { }

        protected override object ParseValue(ICell value) => (int)value.NumericCellValue;
    }
}
