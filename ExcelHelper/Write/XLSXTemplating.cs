using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Linq.Expressions;

namespace ExcelHelper.Write
{
    internal class XLSXTemplating<T> : IXLSXTemplating<T>
    {
        private XLSXWriter<T> _writer;
        public int InsertInd { get; private set; }
        public bool MoveFooter { get; private set; }

        public XLSXTemplating(XLSXWriter<T> writer)
        {
            _writer = writer;
        }

        public IXLSXWriter<T> AddConditionRowStyle(Func<T, string> styleName) =>
            _writer.AddConditionRowStyle(styleName);

        public IXLSXWriter<T> AddRule(Expression<Func<T, object>> propName, int colIndex, string styleName = null) =>
            _writer.AddRule(propName, colIndex, styleName);

        public IXLSXWriter<T> AddStyle(Action<ICellStyle> styling, string styleName) =>
            _writer.AddStyle(styling, styleName);

        public IXLSXTemplating<T> FromEmptyWithHeaders(string[] headers, string styleName = null)
        {
            _writer._workbook = new XSSFWorkbook();
            _writer._sheet = _writer._workbook.CreateSheet();

            var row = _writer._sheet.CreateRow(InsertInd++);
            var style = _writer._styling.GetStyle(styleName);

            for (int i = 0; i < headers.Length; i++)
            {
                var cell = row.CreateCell(i);
                cell.SetCellValue(headers[i]);
                cell.CellStyle = style;
            }

            return this;
        }

        public IXLSXTemplating<T> FromTemplate(string filePath, int page, int insertInd, bool moveFooter)
        {
            _writer._workbook = new XSSFWorkbook(filePath);
            _writer._sheet = _writer._workbook.GetSheetAt(page);
            InsertInd = insertInd;
            MoveFooter = moveFooter;

            return this;
        }

        public IXLSXTemplating<T> ReplaceShortCode(string shortCode, string value)
        {
            foreach (IRow row in _writer._sheet)
            {
                foreach (ICell cell in row)
                {
                    if (cell != null)
                    {
                        var shortcodeKey = $"[[[{shortCode}]]]";
                        if (cell.ToString().Contains(shortcodeKey))
                        {
                            var newValue = cell.ToString().Replace(shortcodeKey, value);
                            cell.SetCellValue(newValue);
                        }
                    }
                }
            }

            return this;
        }

        public void Generate(string resultFilePath, T[] models)
        {
            _writer.Generate(resultFilePath, models);
        }

        public void MoveFooterIfNeed(int len)
        {
            if (MoveFooter && _writer._sheet.LastRowNum > InsertInd)
            {
                var afterHeaderRow = InsertInd + 1;
                var templateLastRow = _writer._sheet.LastRowNum;

                _writer._sheet.ShiftRows(afterHeaderRow, templateLastRow, len);
            }
        }

        public IXLSXWriter<T> Modify(Action<XSSFWorkbook, ISheet> action)
        {
            return _writer.Modify(action);
        }
    }
}
