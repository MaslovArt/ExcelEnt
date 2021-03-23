using NPOI.SS.UserModel;
using System;

namespace ExcelEnt.Write
{
    public interface IXLSXStyling<T>
    {
        IXLSXStyling<T> AddCellsStyle(string styleName);
        IXLSXStyling<T> AddCellsStyle(string styleName, int columnIndex);
        IXLSXStyling<T> AddConditionRowStyle(Func<T, int, string> styleName);
        IXLSXStyling<T> AddStyle(Action<ICellStyle> styling, string styleName);
    }
}