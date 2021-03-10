using ExcelEnt.Extentions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Linq.Expressions;

namespace ExcelEnt.Write
{
    /// <summary>
    /// Entities to excel writer
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public class XLSXWriter<T>
    {
        internal List<WriteRule>    _rules;
        internal XSSFWorkbook       _workbook;
        internal ISheet             _sheet;

        internal XLSXStyling<T>     _styling;
        internal XLSXTemplating<T>  _templating;

        public XLSXWriter()
        {
            _workbook = new XSSFWorkbook();
            _sheet = _workbook.CreateSheet();

            _rules = new List<WriteRule>();
            _templating = new XLSXTemplating<T>();
            _styling = new XLSXStyling<T>(_workbook, _sheet);
        }

        private int InsertIndex => _templating.InsertInd;

        public XLSXWriter<T> AddRule(Expression<Func<T, object>> propName, int colIndex, string styleName = null)
        {
            var prop = TypeExtentions.GetProperty(propName);
            _rules.Add(new WriteRule(colIndex, prop, styleName));

            return this;
        }

        public XLSXWriter<T> FromTemplate(string filePath, int page, int insertInd, bool moveFooter)
        {
            _workbook = _templating.FromTemplate(filePath, insertInd, moveFooter);
            _sheet = _workbook.GetSheetAt(page);
            _styling = new XLSXStyling<T>(_workbook, _sheet);

            return this;
        }

        public XLSXWriter<T> FromEmptyWithHeaders(string[] headers, string styleName = null)
        {
            _templating.FromEmptyWithHeaders(_workbook, headers);

            return this;
        }

        public XLSXWriter<T> ReplaceShortCode(string shortcode, string value)
        {
            _templating.ReplaceShortCode(_sheet, shortcode, value);

            return this;
        }

        public XLSXWriter<T> AddStyle(Action<ICellStyle> styling, string styleName)
        {
            _styling.AddStyle(styling, styleName);

            return this;
        }

        public XLSXWriter<T> AddConditionRowStyle(Func<T, string> styleName)
        {
            _styling.AddConditionRowStyle(styleName);

            return this;
        }

        public XLSXWriter<T> Modify(Action<XSSFWorkbook, ISheet> action)
        {
            action(_workbook, _sheet);

            return this;
        }

        public void Generate(string resultFilePath, T[] entities)
        {
            _templating.MoveFooterIfNeed(_sheet, entities.Length);

            var newRowInd = InsertIndex;
            var minColIndex = _rules.Select(r => r.ExcelColInd).Min();
            var maxColIndex = _rules.Select(r => r.ExcelColInd).Max();

            foreach (var model in entities)
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

            _styling.ApplyConditionRowStyles(entities, InsertIndex);

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
