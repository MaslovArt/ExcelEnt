using NPOI.SS.UserModel;
using System;

namespace ExcelEnt.Write
{
    public interface IXLSXStyling<T>
    {
        IXLSXStyling<T> AddCellsDefaultStyle(string styleName);
        IXLSXStyling<T> AddCellStyle(int cellIndex, string styleName);
        IXLSXStyling<T> AddConditionRowStyle(Func<T, int, string> styleName);
        IXLSXStyling<T> AddStyle(Action<ICellStyle> styling, string styleName);
    }
}