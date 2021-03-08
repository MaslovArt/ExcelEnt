using System.Reflection;

namespace ExcelHelper.Write
{
    internal class WriteRule
    {
        internal WriteRule(int excelColInd, PropertyInfo prop)
        {
            ExcelColInd = excelColInd;
            Prop = prop;
        }

        internal PropertyInfo Prop { get; set; }

        internal int ExcelColInd { get; set; }
    }
}
