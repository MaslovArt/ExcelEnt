using NPOI.SS.UserModel;
using System;
using System.Reflection;

namespace ExcelEnt.Bind
{
    /// <summary>
    /// Excel to entity map rule
    /// </summary>
    internal class BindRule
    {
        /// <summary>
        /// Create excel to entity map rule
        /// </summary>
        /// <param name="excelColInd">Excel file column index</param>
        /// <param name="prop">Entity property</param>
        /// <param name="mapper">Excel cell to entity field mapper</param>
        internal BindRule(int excelColInd, PropertyInfo prop, Func<ICell, object> mapper)
        {
            ExcelColInd = excelColInd;
            Prop = prop;
            Map = mapper;
        }

        /// <summary>
        /// Entity property
        /// </summary>
        internal PropertyInfo Prop { get; set; }

        /// <summary>
        /// Excel file column index
        /// </summary>
        internal int ExcelColInd { get; set; }

        /// <summary>
        /// Excel cell to entity field mapper
        /// </summary>
        internal Func<ICell, object> Map { get; set; }
    }
}
