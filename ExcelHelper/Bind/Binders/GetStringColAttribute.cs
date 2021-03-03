using NPOI.SS.UserModel;

namespace ExcelHelper.Bind.Binders
{
    public class GetStringColAttribute : BaseColAttribute
    {
        public GetStringColAttribute(int columnIndex)
            : base(columnIndex) { }

        protected override object ParseValue(ICell value) => value.ToString();
    }
}
