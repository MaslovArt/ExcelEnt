using NPOI.SS.UserModel;
using System;
using System.Linq.Expressions;

namespace ExcelHelper.Write
{
    public interface IXLSXWriter<T>
    {
        IXLSXWriter<T> AddConditionRowStyle(Func<T, string> styleName);
        IXLSXWriter<T> AddRule(Expression<Func<T, object>> propName, int colIndex, string styleName = null);
        IXLSXWriter<T> AddStyle(Action<ICellStyle> styling, string styleName);
        IXLSXTemplating<T> FromEmptyWithHeaders(string[] headers, string styleName = null);
        IXLSXTemplating<T> FromTemplate(string filePath, int page, int insertInd, bool moveFooter);
        void Generate(string resultFilePath, T[] models);
    }
}