using ExcelHelper.Extentions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
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

        public XLSXWriter()
        {
            CreateWB();
            _rules = new List<WriteRule>();
            _styles = new Dictionary<string, ICellStyle>();
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

        public XLSXWriter<T> FromTemplate(string filePath, int page, int insertInd, bool moveFooter)
        {
            _insertInd = insertInd;
            _moveFooter = moveFooter;
            _workbook = new XSSFWorkbook(filePath);
            _sheet = _workbook.GetSheetAt(page);

            return this;
        }

        public XLSXWriter<T> FromEmptyWithHeaders(string[] headers)
        {
            CreateWB();
            var row = _sheet.CreateRow(_insertInd++);

            for (int i = 0; i < headers.Length; i++)
                row.CreateCell(i).SetCellValue(headers[i]);

            return this;
        }

        public XLSXWriter<T> AddStyle(Action<ICellStyle> styling, string styleName)
        {
            var newStyle = _workbook.CreateCellStyle();
            styling(newStyle);

            _styles.Add(styleName, newStyle);

            return this;
        }

        public void Generate(string resultFilePath, T[] models)
        {
            MoveFooterIfNeed(_moveFooter, _sheet, _insertInd, models.Length);

            var newRowInd = _insertInd;

            foreach (var model in models)
            {
                var row = _sheet.CreateRow(newRowInd++);
                foreach (var rule in _rules)
                {
                    var value = rule.Prop.GetValue(model);
                    var newCell = CreateCell(row, rule.ExcelColInd, rule.StyleName);

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

            SaveExcel(resultFilePath);
        }


        private ICell CreateCell(IRow row, int cellInd, string cellStyleName)
        {
            var newCell = row.CreateCell(cellInd);

            if (!string.IsNullOrEmpty(cellStyleName))
                newCell.CellStyle = _styles[cellStyleName];

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
