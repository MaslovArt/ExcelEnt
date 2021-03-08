using NPOI.SS.UserModel;
using System;
using System.Reflection;

namespace ExcelHelper.Bind
{
    internal class BindRule
    {
        internal BindRule(int excelColInd, PropertyInfo prop, Func<ICell, object> mapper)
        {
            ExcelColInd = excelColInd;
            Prop = prop;
            Map = mapper;
        }

        internal PropertyInfo Prop { get; set; }

        internal int ExcelColInd { get; set; }

        internal Func<ICell, object> Map { get; set; }
    }
}
