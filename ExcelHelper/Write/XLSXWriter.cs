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
        private List<WriteRule> Rules;

        public XSSFWorkbook Workbook { get; private set; }

        public ISheet Sheet { get; private set; }

        public string TemplateFilePath { get; private set; }

        public int InsertInd { get; private set; }

        public bool MoveFooter { get; private set; }

        public XLSXWriter()
        {
            Workbook = new XSSFWorkbook();
            Sheet = Workbook.CreateSheet();
            Rules = new List<WriteRule>();
        }

        public XLSXWriter<T> AddRule(Expression<Func<T, object>> propName, int colIndex)
        {
            var prop = TypeExtentions.GetProperty(propName);
            Rules.Add(new WriteRule(colIndex, prop));

            return this;
        }

        public XLSXWriter<T> FromTemplate(string filePath, int page, int insertInd, bool moveFooter)
        {
            TemplateFilePath = filePath;
            InsertInd = insertInd;
            MoveFooter = moveFooter;
            Workbook = new XSSFWorkbook(filePath);
            Sheet = Workbook.GetSheetAt(page);

            return this;
        }

        public void Generate(string resultFilePath, T[] models)
        {
            MoveFooterIfNeed(MoveFooter, Sheet, InsertInd, models.Length);

            var newRowInd = InsertInd;

            foreach (var model in models)
            {
                var row = Sheet.CreateRow(newRowInd++);
                foreach (var rule in Rules)
                {
                    var value = rule.Prop.GetValue(model);
                    var newCell = row.CreateCell(rule.ExcelColInd);

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
                Workbook.Write(file, true);
            }
        }
    }
}
