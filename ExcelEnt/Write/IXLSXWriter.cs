using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Linq.Expressions;

namespace ExcelEnt.Write
{
    public interface IXLSXWriter<T>
    {
        /// <summary>
        /// Add row styling by condition
        /// </summary>
        /// <param name="styleName">Existing style name</param>
        /// <returns></returns>
        IXLSXWriter<T> AddConditionRowStyle(Func<T, string> styleName);

        /// <summary>
        /// Add entity property value to cell value rule
        /// </summary>
        /// <param name="propName">Entity property name</param>
        /// <param name="colIndex">Excel cell index</param>
        /// <param name="styleName">Existing style name</param>
        /// <returns></returns>
        IXLSXWriter<T> AddRule(Expression<Func<T, object>> propName, int colIndex, string styleName = null);

        /// <summary>
        /// Add style to current workbook
        /// </summary>
        /// <param name="styling">Style</param>
        /// <param name="styleName">Style name</param>
        /// <returns></returns>
        IXLSXWriter<T> AddStyle(Action<ICellStyle> styling, string styleName);

        /// <summary>
        /// Create new workbook with columns titles
        /// </summary>
        /// <param name="headers">Columns titles</param>
        /// <param name="styleName">Titles existing style name</param>
        /// <returns></returns>
        IXLSXTemplating<T> FromEmptyWithHeaders(string[] headers, string styleName = null);

        /// <summary>
        /// Create new workbook from excel template
        /// </summary>
        /// <param name="filePath">Excel template file path</param>
        /// <param name="page">Excel sheet index</param>
        /// <param name="insertInd">Inserting index</param>
        /// <param name="moveFooter">Move all rows after insert index</param>
        /// <returns></returns>
        IXLSXTemplating<T> FromTemplate(string filePath, int page, int insertInd, bool moveFooter);

        /// <summary>
        /// Modidy current workbook
        /// </summary>
        /// <param name="action">Modification</param>
        /// <returns></returns>
        IXLSXWriter<T> Modify(Action<XSSFWorkbook, ISheet> action);

        /// <summary>
        /// Generate new excel file
        /// </summary>
        /// <param name="resultFilePath">Generated file path</param>
        /// <param name="entities">Entities for write</param>
        void Generate(string resultFilePath, T[] entities);
    }
}