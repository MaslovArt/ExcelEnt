using NPOI.SS.UserModel;
using System;
using System.ComponentModel;

namespace ExcelHelper.Bind.Binders
{
    public class EnumColAttribute : BaseColAttribute
    {
        public EnumColAttribute(int columnIndex, Type enumType)
            : base(columnIndex)
        {
            EnumType = enumType;
        }

        public Type EnumType { get; set; }

        protected override object ParseValue(ICell value)
        {
            var description = value.StringCellValue;

            if (!EnumType.IsEnum)
                throw new ArgumentException();

            foreach (var field in EnumType.GetFields())
            {
                var attribute = GetCustomAttribute(field, typeof(DescriptionAttribute)) as DescriptionAttribute;
                if (attribute == null)
                    continue;
                if (attribute.Description == description)
                {
                    return (int)field.GetValue(null);
                }
            }

            throw new ArgumentException(description);
        }
    }
}
