using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;

namespace ExcelEnt.Write
{
    /// <summary>
    /// Excel styling
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public class XLSXStyling<T>
    {
        private Dictionary<string, ICellStyle>  _styles;
        private List<Func<T, string>>           _conditionStyles;
        private string                          _rowsDefaultStyleName;
        private ICellStyle                      _defaultStyle;

        private XSSFWorkbook                    _workbook;
        private ISheet                          _sheet;

        public XLSXStyling(XSSFWorkbook workbook, ISheet sheet)
        {
            _workbook = workbook;
            _sheet = sheet;
            _styles = new Dictionary<string, ICellStyle>();
            _conditionStyles = new List<Func<T, string>>();
        }

        /// <summary>
        /// Add style to current sheet
        /// </summary>
        /// <param name="styling">Style definition</param>
        /// <param name="styleName">Style name</param>
        public XLSXStyling<T> AddStyle(Action<ICellStyle> styling, string styleName)
        {
            var newStyle = _workbook.CreateCellStyle();
            styling(newStyle);

            if (_defaultStyle == null && styleName == _rowsDefaultStyleName)
            {
                _defaultStyle = newStyle;
            }

            _styles.Add(styleName, newStyle);

            return this;
        }

        /// <summary>
        /// Add row style by condition
        /// </summary>
        /// <param name="styleName"></param>
        public XLSXStyling<T> AddConditionRowStyle(Func<T, string> styleName)
        {
            _conditionStyles.Add(styleName);

            return this;
        }

        /// <summary>
        /// Add default rows styling
        /// </summary>
        /// <param name="styleName">Existing style name</param>
        public XLSXStyling<T> AddRowDefaultStyle(string styleName)
        {
            _rowsDefaultStyleName = styleName;
            _defaultStyle = _styles[styleName];

            return this;
        } 

        /// <summary>
        /// Apply condition row styles
        /// </summary>
        /// <param name="entities">Entities</param>
        /// <param name="insertIndex">Entities insert index</param>
        internal void ApplyConditionRowStyles(T[] entities, int insertIndex)
        {
            if (_conditionStyles.Count == 0) return;

            for (int i = 0; i < entities.Length; i++)
            {
                foreach (var condition in _conditionStyles)
                {
                    var styleName = condition(entities[i]);
                    if (!string.IsNullOrEmpty(styleName))
                    {
                        var row = _sheet.GetRow(i + insertIndex);
                        var style = _styles[styleName];

                        foreach (var cell in row)
                            cell.CellStyle = style;
                    }
                }
            }
        }

        /// <summary>
        /// Set cell style by name or row default
        /// </summary>
        /// <param name="cell">Cell</param>
        /// <param name="name">Style name</param>
        /// <returns></returns>
        internal void SetStyle(ICell cell, string name)
        {
            ICellStyle style = null;
            if (!string.IsNullOrEmpty(name))
            {
                style = _styles[name];
            }

            cell.CellStyle = style ?? _defaultStyle;
        }
    }
}
