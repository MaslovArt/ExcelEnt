using NPOI.SS.UserModel;

namespace ExcelHelper.Bind.Binders
{
    public class ReadStringColAttribute : BaseColAttribute
    {
        public ReadStringColAttribute(int columnIndex)
            : base(columnIndex) { }

        protected override object ParseValue(ICell value) => value.ToString();
    }
}
