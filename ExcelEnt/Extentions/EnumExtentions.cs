using System;
using System.ComponentModel;
using System.Reflection;

namespace ExcelEnt.Extentions
{
    internal static class EnumExtentions
    {
        internal static string ToDescription<T>(this T _enum) =>
            _enum.GetType().GetField(_enum.ToString()).DescriptionAttrValue();

        internal static Enum ToEnum<Enum>(this string description)
        {
            var type = typeof(Enum);
            if (!type.IsEnum)
                throw new ArgumentException();

            foreach (var field in type.GetFields())
            {
                var attribute = field.GetCustomAttribute<DescriptionAttribute>();
                if (attribute == null)
                    continue;
                if (attribute.Description == description)
                {
                    return (Enum)field.GetValue(null);
                }
            }

            throw new ArgumentException(description);
        }
    }
}
