using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Collections.Generic;

namespace ExcelEnt.Write
{
    /// <summary>
    /// Excel templating
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public class XLSXTemplating<T>
    {
        private Dictionary<string, string> _shortcodes;
        internal bool _moveFooter;
        internal string[] _headers;
        internal string _templatePath { get; private set; }

        internal int InsertInd { get; private set; }

        public XLSXTemplating()
        {
            _shortcodes = new Dictionary<string, string>();
        }

        public XLSXTemplating<T> FromEmptyWithHeaders(string[] headers)
        {
            InsertInd = 0;
            _headers = headers;

            return this;
        }

        public XLSXTemplating<T> FromTemplate(string filePath, int insertInd, bool moveFooter)
        {
            InsertInd = insertInd;
            _moveFooter = moveFooter;
            _templatePath = filePath;

            return this;
        }

        public XLSXTemplating<T> ReplaceShortCode(string shortCode, string value)
        {
            _shortcodes.Add(shortCode, value);

            return this;
        }

        public XSSFWorkbook CreateWorkbook(int entitiesCount)
        {
            var workbook = _templatePath != null
                ? CreateFromTemplate()
                : CreateFromEmptyWithHeaders();

            var sheet = workbook.GetSheetAt(0);

            ApplyShortcodesReplace(sheet);
            MoveFooterIfNeed(sheet, entitiesCount);

            return workbook;
        }

        private XSSFWorkbook CreateFromEmptyWithHeaders()
        {
            var workbook = new XSSFWorkbook();
            var sheet = workbook.GetSheetAt(0);

            if (_headers != null && _headers.Length > 0)
            {
                var row = sheet.CreateRow(InsertInd++);
                for (int i = 0; i < _headers.Length; i++)
                {
                    var cell = row.CreateCell(i);
                    cell.SetCellValue(_headers[i]);
                }
            }

            return workbook;
        }

        private XSSFWorkbook CreateFromTemplate()
        {
            return new XSSFWorkbook(_templatePath);
        }

        private void ApplyShortcodesReplace(ISheet sheet)
        {
            foreach (IRow row in sheet)
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
                                var newValue = cell.ToString().Replace(shortcodeKey, shortcode.Value);
                                cell.SetCellValue(newValue);
                            }
                        }
                    }
                }
            }
        }

        private void MoveFooterIfNeed(ISheet sheet, int len)
        {
            if (_moveFooter && sheet.LastRowNum > InsertInd)
            {
                var afterHeaderRow = InsertInd + 1;
                var templateLastRow = sheet.LastRowNum;

                sheet.ShiftRows(afterHeaderRow, templateLastRow, len);
            }
        }
    }
}
