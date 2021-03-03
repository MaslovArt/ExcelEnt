using System;
using System.Reflection;

namespace ExcelHelper
{
    internal class BindProp<T> where T : Attribute
    {
        internal BindProp(PropertyInfo prop, T attribute)
        {
            Prop = prop;
            Attribute = attribute;
        }

        internal PropertyInfo Prop { get; set; }
        internal T Attribute { get; set; }
    }
}
