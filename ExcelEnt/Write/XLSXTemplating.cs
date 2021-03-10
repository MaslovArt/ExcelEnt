using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace ExcelEnt.Write
{
    /// <summary>
    /// Excel templating
    /// </summary>
    /// <typeparam name="T"></typeparam>
    internal class XLSXTemplating<T>
    {
        internal int InsertInd { get; private set; }
        internal bool MoveFooter { get; private set; }

        /// <summary>
        /// Create workbook with columns titles
        /// </summary>
        /// <param name="workbook">Excel workbook</param>
        /// <param name="headers">Columns titles</param>
        internal void FromEmptyWithHeaders(XSSFWorkbook workbook, string[] headers)
        {
            InsertInd = 0;
            var sheet = workbook.GetSheetAt(0);
            var row = sheet.CreateRow(InsertInd++);

            for (int i = 0; i < headers.Length; i++)
            {
                var cell = row.CreateCell(i);
                cell.SetCellValue(headers[i]);
            }
        }

        /// <summary>
        /// Create excel from template
        /// </summary>
        /// <param name="filePath">Excel template path</param>
        /// <param name="insertInd">Entities insert index</param>
        /// <param name="moveFooter">Move cells after inserting index</param>
        /// <returns></returns>
        internal XSSFWorkbook FromTemplate(string filePath, int insertInd, bool moveFooter)
        {
            var workbook = new XSSFWorkbook(filePath);
            InsertInd = insertInd;
            MoveFooter = moveFooter;

            return workbook;
        }

        /// <summary>
        /// Replace template shortcodes with value
        /// </summary>
        /// <param name="sheet">Excel sheet</param>
        /// <param name="shortCode">Shortcode</param>
        /// <param name="value">Shortcode value</param>
        internal void ReplaceShortCode(ISheet sheet, string shortCode, string value)
        {
            foreach (IRow row in sheet)
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
        }

        /// <summary>
        /// Move footer row for entities inserting
        /// </summary>
        /// <param name="sheet">Excel sheet</param>
        /// <param name="len">Offset</param>
        internal void MoveFooterIfNeed(ISheet sheet, int len)
        {
            if (MoveFooter && sheet.LastRowNum > InsertInd)
            {
                var afterHeaderRow = InsertInd + 1;
                var templateLastRow = sheet.LastRowNum;

                sheet.ShiftRows(afterHeaderRow, templateLastRow, len);
            }
        }
    }
}
