using System;

namespace ExcelHelper.Write
{
    [AttributeUsage(AttributeTargets.Field | AttributeTargets.Property)]
    public class WriteColAttribute : Attribute
    {
        public WriteColAttribute(int columnIndex)
        {
            ColumnIndex = columnIndex;
        }

        public int ColumnIndex { get; set; }

        public string StyleName { get; set; }
    }
}
