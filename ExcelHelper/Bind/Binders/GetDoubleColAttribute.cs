using NPOI.SS.UserModel;

namespace ExcelHelper.Bind.Binders
{
    public class GetDoubleColAttribute : BaseColAttribute
    {
        public GetDoubleColAttribute(int columnIndex)
            : base(columnIndex) { }

        protected override object ParseValue(ICell value) => value.NumericCellValue;
    }
}
