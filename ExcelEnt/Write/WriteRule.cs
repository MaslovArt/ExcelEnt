using System.Reflection;

namespace ExcelEnt.Write
{
    /// <summary>
    /// Entity to excel write rule
    /// </summary>
    internal class WriteRule
    {
        internal WriteRule(int excelColInd, PropertyInfo prop)
        {
            ExcelColInd = excelColInd;
            Prop = prop;
        }

        /// <summary>
        /// Entity property
        /// </summary>
        internal PropertyInfo Prop { get; set; }

        /// <summary>
        /// Excel cell index
        /// </summary>
        internal int ExcelColInd { get; set; }
    }
}
