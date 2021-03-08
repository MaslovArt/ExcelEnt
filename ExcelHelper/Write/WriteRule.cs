using System.Reflection;

namespace ExcelHelper.Write
{
    internal class WriteRule
    {
        internal WriteRule(int excelColInd, PropertyInfo prop, string styleName)
        {
            ExcelColInd = excelColInd;
            Prop = prop;
            StyleName = styleName;
        }

        internal PropertyInfo Prop { get; set; }
        internal int ExcelColInd { get; set; }
        internal string StyleName { get; set; }
    }
}
