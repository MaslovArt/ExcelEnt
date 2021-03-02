using NPOI.SS.UserModel;

namespace ExcelHelper.Binders
{
    public class StringColAttribute : BaseColAttribute
    {
        public StringColAttribute(int columnIndex)
            : base(columnIndex) { }

        protected override object ParseValue(ICell value) => value.ToString();
    }
}
