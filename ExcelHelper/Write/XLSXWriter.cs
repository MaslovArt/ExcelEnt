using ExcelHelper.Extentions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.IO;

namespace ExcelHelper.Write
{
    public class XLSXWriter<T>
    {
        public XSSFWorkbook Workbook { get; private set; }
        public ISheet Sheet { get; private set; }
        public string TemplateFilePath { get; private set; }
        public int InsertInd { get; private set; }
        public bool MoveFooter { get; private set; }

        public XLSXWriter<T> UseTemplate(string filePath, int page, int insertInd, bool moveFooter)
        {
            TemplateFilePath = filePath;
            InsertInd = insertInd;
            MoveFooter = moveFooter;
            Workbook = new XSSFWorkbook(filePath);
            Sheet = Workbook.GetSheetAt(page);

            return this;
        }

        public XLSXWriter<T> UseModelDescription()
        {
            EnsureWorkbook();

            var rules = typeof(T).BindPropsAttrs<WriteColAttribute>();
            var row = Sheet.CreateRow(InsertInd++);

            foreach (var rule in rules)
            {
                var title = rule.Prop.DescriptionAttrValue();
                row.CreateCell(rule.Attribute.ColumnIndex, CellType.String).SetCellValue(title);
            }

            return this;
        }

        public void Generate(string resultFilePath, T[] models)
        {
            EnsureWorkbook();
            MoveFooterIfNeed(MoveFooter, Sheet, InsertInd, models.Length);

            var newRowInd = InsertInd;
            var rules = typeof(T).BindPropsAttrs<WriteColAttribute>();

            foreach (var model in models)
            {
                var row = Sheet.CreateRow(newRowInd++);
                foreach (var rule in rules)
                {
                    var value = rule.Prop.GetValue(model);
                    var newCell = row.CreateCell(rule.Attribute.ColumnIndex);

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

        private void EnsureWorkbook()
        {
            Workbook = Workbook ?? new XSSFWorkbook();
            Sheet = Sheet ?? Workbook.CreateSheet();
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
