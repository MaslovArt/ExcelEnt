using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;

namespace ExcelHelper.Write
{
    internal class XLSXStyling<T>
    {
        private Dictionary<string, ICellStyle> _styles;
        private List<Func<T, string>> _conditionStyles;

        private XSSFWorkbook _workbook;
        private ISheet _sheet;

        public XLSXStyling(XSSFWorkbook workbook, ISheet sheet)
        {
            _workbook = workbook;
            _sheet = sheet;
            _styles = new Dictionary<string, ICellStyle>();
            _conditionStyles = new List<Func<T, string>>();
        }

        public void AddStyle(Action<ICellStyle> styling, string styleName)
        {
            var newStyle = _workbook.CreateCellStyle();
            styling(newStyle);

            _styles.Add(styleName, newStyle);
        }

        public void AddConditionRowStyle(Func<T, string> styleName)
        {
            _conditionStyles.Add(styleName);
        }

        public void ApplyConditionRowStyles(T[] models, int insertIndex)
        {
            if (_conditionStyles.Count == 0) return;

            for (int i = 0; i < models.Length; i++)
            {
                foreach (var condition in _conditionStyles)
                {
                    var styleName = condition(models[i]);
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

        public ICellStyle GetStyle(string name)
        {
            if (!string.IsNullOrEmpty(name))
                return _styles[name];

            return null;
        }
    }
}
