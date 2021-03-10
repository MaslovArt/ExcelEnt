using System.Reflection;

namespace ExcelHelper.Write
{
    /// <summary>
    /// Entity to excel write rule
    /// </summary>
    internal class WriteRule
    {
        internal WriteRule(int excelColInd, PropertyInfo prop, string styleName)
        {
            ExcelColInd = excelColInd;
            Prop = prop;
            StyleName = styleName;
        }

        /// <summary>
        /// Entity property
        /// </summary>
        internal PropertyInfo Prop { get; set; }

        /// <summary>
        /// Excel cell index
        /// </summary>
        internal int ExcelColInd { get; set; }

        /// <summary>
        /// Cell style name
        /// </summary>
        internal string StyleName { get; set; }
    }
}
