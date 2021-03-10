using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Linq.Expressions;

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

        internal XSSFWorkbook FromTemplate(string filePath, int insertInd, bool moveFooter)
        {
            var workbook = new XSSFWorkbook(filePath);
            InsertInd = insertInd;
            MoveFooter = moveFooter;

            return workbook;
        }

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
