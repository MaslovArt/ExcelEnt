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
    public class XLSXWriter<T> : IXLSXWriter<T>
    {
        internal List<WriteRule> _rules;
        internal XSSFWorkbook _workbook;
        internal ISheet _sheet;

        internal XLSXStyling<T> _styling;
        internal XLSXTemplating<T> _templating;

        public XLSXWriter()
        {
            _workbook = new XSSFWorkbook();
            _sheet = _workbook.CreateSheet();

            _rules = new List<WriteRule>();
            _styling = new XLSXStyling<T>(_workbook, _sheet);
        }

        private int InsertIndex => _templating?.InsertInd ?? 0;

        public IXLSXWriter<T> AddRule(Expression<Func<T, object>> propName, int colIndex, string styleName = null)
        {
            var prop = TypeExtentions.GetProperty(propName);
            _rules.Add(new WriteRule(colIndex, prop, styleName));

            return this;
        }

        public IXLSXTemplating<T> FromTemplate(string filePath, int page, int insertInd, bool moveFooter)
        {
            _templating = new XLSXTemplating<T>(this);
            _templating.FromTemplate(filePath, page, insertInd, moveFooter);

            _styling = new XLSXStyling<T>(_workbook, _sheet);

            return _templating;
        }

        public IXLSXTemplating<T> FromEmptyWithHeaders(string[] headers, string styleName = null)
        {
            _templating = new XLSXTemplating<T>(this);
            _templating.FromEmptyWithHeaders(headers, styleName);

            _styling = new XLSXStyling<T>(_workbook, _sheet);

            return _templating;
        }

        public IXLSXWriter<T> AddStyle(Action<ICellStyle> styling, string styleName)
        {
            _styling.AddStyle(styling, styleName);

            return this;
        }

        public IXLSXWriter<T> AddConditionRowStyle(Func<T, string> styleName)
        {
            _styling.AddConditionRowStyle(styleName);

            return this;
        }

        public IXLSXWriter<T> Modify(Action<XSSFWorkbook, ISheet> action)
        {
            action(_workbook, _sheet);

            return this;
        }

        public void Generate(string resultFilePath, T[] models)
        {
            _templating?.MoveFooterIfNeed(models.Length);

            var newRowInd = InsertIndex;
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

            _styling.ApplyConditionRowStyles(models, InsertIndex);

            SaveExcel(resultFilePath);
        }


        private ICell GetCell(IRow row, WriteRule rule)
        {
            var newCell = row.GetCell(rule.ExcelColInd);
            newCell.CellStyle = _styling.GetStyle(rule.StyleName);

            return newCell;
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
