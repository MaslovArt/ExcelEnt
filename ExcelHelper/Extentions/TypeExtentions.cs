using System;
using System.ComponentModel;
using System.Linq;
using System.Reflection;

namespace ExcelHelper.Extentions
{
    internal static class TypeExtentions
    {
        internal static BindProp<T>[] BindPropsAttrs<T>(this Type type) where T : Attribute
        {
            return type
                .GetProperties()
                .Select(p => new BindProp<T>(p, p.GetCustomAttribute(typeof(T), true) as T))
                .Where(a => a.Attribute != null)
                .ToArray();
        }

        internal static string DescriptionAttrValue(this MemberInfo propertyInfo)
        {
            var attributes = (DescriptionAttribute[])propertyInfo.GetCustomAttributes(
                typeof(DescriptionAttribute), false);

            if (attributes != null && attributes.Length > 0)
                return attributes[0].Description;

            return null;
        }
    }
}
