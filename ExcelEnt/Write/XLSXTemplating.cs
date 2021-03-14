using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Collections.Generic;

namespace ExcelEnt.Write
{
    /// <summary>
    /// Excel templating
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public class XLSXTemplating<T> : IXLSXTemplating<T>
    {
        private Dictionary<string, string> _shortcodes;
        private bool _moveFooter;
        private string _templatePath;

        internal int InsertInd { get; private set; }

        public XLSXTemplating()
        {
            _shortcodes = new Dictionary<string, string>();
        }

        /// <summary>
        /// Workbook from template
        /// </summary>
        /// <param name="filePath">Template path</param>
        /// <param name="insertInd">Template entities insertion index</param>
        /// <param name="moveFooter">Move rows after insert index</param>
        /// <returns></returns>
        public IXLSXTemplating<T> FromTemplate(string filePath, int insertInd, bool moveFooter)
        {
            InsertInd = insertInd;
            _moveFooter = moveFooter;
            _templatePath = filePath;

            return this;
        }

        /// <summary>
        /// Replace shortcode
        /// </summary>
        /// <param name="shortCode"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        public IXLSXTemplating<T> ReplaceShortCode(string shortCode, string value)
        {
            _shortcodes.Add(shortCode, value);

            return this;
        }

        /// <summary>
        /// Create new workbook
        /// </summary>
        /// <param name="entitiesCount">Entities count (for moving footer)</param>
        /// <returns></returns>
        public XSSFWorkbook CreateWorkbook(int entitiesCount)
        {
            var workbook = new XSSFWorkbook(_templatePath);
            var sheet = workbook.GetSheetAt(0);

            ApplyShortcodesReplace(sheet);
            MoveFooterIfNeed(sheet, entitiesCount);

            return workbook;
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
