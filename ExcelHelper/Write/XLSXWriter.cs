using ExcelHelper.Extentions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Linq.Expressions;

namespace ExcelHelper.Write
{
    public class XLSXWriter<T>
    {
        private List<WriteRule> _rules;
        private XSSFWorkbook _workbook;
        private ISheet _sheet;
        private int _insertInd;
        private bool _moveFooter;

        private Dictionary<string, ICellStyle> _styles;
        private Dictionary<string, string> _shortcodes;
        private List<Func<T, string>> _conditionStyles;

        public XLSXWriter()
        {
            CreateWB();
            _rules = new List<WriteRule>();
            _styles = new Dictionary<string, ICellStyle>();
            _shortcodes = new Dictionary<string, string>();
            _conditionStyles = new List<Func<T, string>>();
        }

        public XLSXWriter<T> AddRule(
            Expression<Func<T, object>> propName,
            int colIndex,
            string styleName = null)
        {
            var prop = TypeExtentions.GetProperty(propName);
            _rules.Add(new WriteRule(colIndex, prop, styleName));

            return this;
        }

        public XLSXWriter<T> FromTemplate(
            string filePath, 
            int page, 
            int insertInd, 
            bool moveFooter)
        {
            _insertInd = insertInd;
            _moveFooter = moveFooter;
            _workbook = new XSSFWorkbook(filePath);
            _sheet = _workbook.GetSheetAt(page);

            return this;
        }

        public XLSXWriter<T> ReplaceShortCode(string shortCode, string value)
        {
            _shortcodes.Add(shortCode, value);

            return this;
        }

        public XLSXWriter<T> FromEmptyWithHeaders(string[] headers, string styleName = null)
        {
            CreateWB();
            var row = _sheet.CreateRow(_insertInd++);
            var style = !string.IsNullOrEmpty(styleName) ? _styles[styleName] : null;

            for (int i = 0; i < headers.Length; i++)
            {
                var cell = row.CreateCell(i);
                cell.SetCellValue(headers[i]);
                cell.CellStyle = style;
            }

            return this;
        }

        public XLSXWriter<T> AddStyle(Action<ICellStyle> styling, string styleName)
        {
            var newStyle = _workbook.CreateCellStyle();
            styling(newStyle);

            _styles.Add(styleName, newStyle);

            return this;
        }

        public XLSXWriter<T> AddConditionRowStyle(Func<T, string> styleName)
        {
            _conditionStyles.Add(styleName);

            return this;
        }

        public void Generate(string resultFilePath, T[] models)
        {
            ApplyReplaceShortCodes();
            MoveFooterIfNeed(_moveFooter, _sheet, _insertInd, models.Length);

            var newRowInd = _insertInd;
            var minColIndex = _rules.Select(r => r.ExcelColInd).Min();
            var maxColIndex = _rules.Select(r => r.ExcelColInd).Max();

            foreach (var model in models)
            {
                var row = _sheet.CreateRow(newRowInd++);
                for (var colInd = minColIndex; colInd <= maxColIndex; colInd++)
                    row.CreateCell(colInd);

                foreach (var rule in _rules)
                {
                    var value = rule.Prop.GetValue(model);
                    var newCell = GetCell(row, rule);

                    if (value == null)
                        newCell.SetCellValue("");
                    else if (value is string strValue)
                        newCell.SetCellValue(strValue);
                    else if (value is DateTime dateValue)
                        newCell.SetCellValue(dateValue);
                    else if (value is bool boolValue)
                        newCell.SetCellValue(boolValue);
                    else if (double.TryParse(value.ToString(), out double numValue))
                        newCell.SetCellValue(numValue);
                    else if (value is Enum enumValue)
                        newCell.SetCellValue(enumValue.ToDescription());
                    else
                        newCell.SetCellValue(value.ToString());
                }
            }

            ApplyConditionRowStyles(models);
            SaveExcel(resultFilePath);
        }


        private void ApplyConditionRowStyles(T[] models)
        {
            if (_conditionStyles.Count == 0) return;

            var indexes = _rules.Select(r => r.ExcelColInd);
            var startCellInd = indexes.Min();
            var endCellInd = indexes.Max();

            for (int i = 0; i < models.Length; i++)
            {
                foreach (var cond in _conditionStyles)
                {
                    var styleName = cond(models[i]);
                    if (!string.IsNullOrEmpty(styleName))
                    {
                        var row = _sheet.GetRow(i + _insertInd);
                        var style = _styles[styleName];
                        for (int j = startCellInd; j <= endCellInd; j++)
                            row.GetCell(j).CellStyle = style;
                    }
                }
            }
        }

        private void ApplyReplaceShortCodes()
        {
            if (_shortcodes.Count == 0) return;

            foreach (IRow row in _sheet)
            {
                foreach (ICell cell in row)
                {
                    if (cell != null)
                    {
                        foreach (var shortcode in _shortcodes)
                        {
                            var shortcodeKey = $"[[[{shortcode.Key}]]]";
                            if (cell.ToString().Contains(shortcodeKey))
                            {
                                var newValue = cell.ToString()
                                    .Replace(shortcodeKey, shortcode.Value);
                                cell.SetCellValue(newValue);
                                break;
                            }
                        }
                    }
                }
            }
        }

        private ICell GetCell(IRow row, WriteRule rule)
        {
            var newCell = row.GetCell(rule.ExcelColInd);

            if (!string.IsNullOrEmpty(rule.StyleName))
                newCell.CellStyle = _styles[rule.StyleName];

            return newCell;
        }

        private void CreateWB()
        {
            _workbook = new XSSFWorkbook();
            _sheet = _workbook.CreateSheet();
        }

        private void MoveFooterIfNeed(bool moveFooter, ISheet sheet, int toInd, int len)
        {
            if (moveFooter && sheet.LastRowNum > toInd)
            {
                var afterHeaderRow = toInd + 1;
                var templateLastRow = sheet.LastRowNum;

                sheet.ShiftRows(afterHeaderRow, templateLastRow, len);
            }
        }

        private void SaveExcel(string resultFilePath)
        {
            using (var file = new FileStream(resultFilePath, FileMode.CreateNew))
            {
                _workbook.Write(file, true);
            }
        }
    }
}
