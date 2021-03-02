using NPOI.SS.UserModel;
using System;

namespace ExcelHelper.Bind.Binders
{
    [AttributeUsage(AttributeTargets.Field | AttributeTargets.Property)]
    public abstract class BaseColAttribute : Attribute
    {
        public BaseColAttribute(int columnIndex)
        {
            ColumnIndex = columnIndex;
        }

        public int ColumnIndex { get; set; }

        protected abstract object ParseValue(ICell value);

        public void MapProp(string propName, ICell propValue, object obj)
        {
            var value = ParseValue(propValue);
            obj.GetType().GetProperty(propName).SetValue(obj, value);
        }
    }
}
