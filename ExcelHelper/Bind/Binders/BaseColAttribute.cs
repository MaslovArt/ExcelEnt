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

        public void MapProp(BindProp<BaseColAttribute> bindInfo, ICell propValue, object obj)
        {
            var value = propValue == null 
                ? null 
                : ParseValue(propValue);

            if (value == null && !bindInfo.CanBeNull) 
                throw new Exception($"Set null value to '{bindInfo.Prop.Name}' ({bindInfo.Prop.PropertyType}).");

            bindInfo.Prop.SetValue(obj, value);
        }
    }
}
