using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelEnt.Write
{
    /// <summary>
    /// Excel styling
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public class XLSXStyling<T> : IXLSXStyling<T>
    {
        private Dictionary<string, Action<ICellStyle>> _styleDefinitions;
        private Dictionary<string, ICellStyle> _styles;
        private Dictionary<int, string> _cellsStyles;
        private List<Func<T, int, string>> _conditionStyles;
        private string _cellDefaultStyleName;
        private ICellStyle _defaultStyle;

        public XLSXStyling()
        {
            _styleDefinitions = new Dictionary<string, Action<ICellStyle>>();
            _conditionStyles = new List<Func<T, int, string>>();
            _cellsStyles = new Dictionary<int, string>();
        }

        /// <summary>
        /// Add new style
        /// </summary>
        /// <param name="styling">Definition</param>
        /// <param name="styleName">Style name</param>
        /// <returns></returns>
        public IXLSXStyling<T> AddStyle(Action<ICellStyle> styling, string styleName)
        {
            _styleDefinitions.Add(styleName, styling);

            return this;
        }

        /// <summary>
        /// Add condition row style
        /// </summary>
        /// <param name="styleName">Style func (entite, rowIndex, styleName)</param>
        /// <returns></returns>
        public IXLSXStyling<T> AddConditionRowStyle(Func<T, int, string> styleName)
        {
            _conditionStyles.Add(styleName);

            return this;
        }

        /// <summary>
        /// Add default style for cells
        /// </summary>
        /// <param name="styleName">Style name</param>
        /// <returns></returns>
        public IXLSXStyling<T> AddCellsDefaultStyle(string styleName)
        {
            _cellDefaultStyleName = styleName;

            return this;
        }

        /// <summary>
        /// Add style for column
        /// </summary>
        /// <param name="columnIndex">Column index</param>
        /// <param name="styleName">Style name</param>
        /// <returns></returns>
        public IXLSXStyling<T> AddCellStyle(int columnIndex, string styleName)
        {
            _cellsStyles.Add(columnIndex, styleName);

            return this;
        }

        /// <summary>
        /// Apply style configuration
        /// </summary>
        /// <param name="workbook">Excel workbook</param>
        /// <returns></returns>
        public XLSXStyling<T> Build(XSSFWorkbook workbook)
        {
            _styles = _styleDefinitions.ToDictionary(
                sd => sd.Key,
                sd =>
                {
                    var newStyle = workbook.CreateCellStyle();
                    sd.Value(newStyle);

                    return newStyle;
                });

            _defaultStyle = !string.IsNullOrEmpty(_cellDefaultStyleName)
                ? _styles[_cellDefaultStyleName]
                : null;

            return this;
        }

        public void ApplyConditionRowStyles(ISheet sheet, T[] entities, int insertIndex)
        {
            if (_conditionStyles.Count == 0) return;

            for (int i = 0; i < entities.Length; i++)
            {
                foreach (var condition in _conditionStyles)
                {
                    var styleName = condition(entities[i], i);
                    if (!string.IsNullOrEmpty(styleName))
                    {
                        var row = sheet.GetRow(i + insertIndex);
                        var style = _styles[styleName];

                        foreach (var cell in row)
                            cell.CellStyle = style;
                    }
                }
            }
        }

        public void SetStyle(ICell cell)
        {
            ICellStyle cellStyle = null;

            if (_cellsStyles.TryGetValue(cell.ColumnIndex, out string cellStyleName))
            {
                cellStyle = _styles[cellStyleName];
            }

            cell.CellStyle = cellStyle ?? _defaultStyle;
        }
    }
}
