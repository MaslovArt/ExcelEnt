using System;
using System.Reflection;

namespace ExcelHelper
{
    public class BindProp<T> where T : Attribute
    {
        public BindProp(PropertyInfo prop, T attribute)
        {
            Prop = prop;
            Attribute = attribute;

            var type = prop.PropertyType;
            CanBeNull = !type.IsValueType || (Nullable.GetUnderlyingType(type) != null);
        }

        public PropertyInfo Prop { get; set; }
        public T Attribute { get; set; }
        public bool CanBeNull { get; set; }
    }
}
