using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Reflection;

namespace ExcelHelper.Write
{
    public class XLSXWriter
    {
        public void Generate<T>(string fromTemplatePath, string sheetName, string toFile, int toInd, T[] models, bool moveFooter = false)
        {
            var workbook = new XSSFWorkbook(fromTemplatePath);
            var sheet = workbook.GetSheet(sheetName);
            var newRowInd = toInd;

            MoveFooterIfNeed(moveFooter, sheet, toInd, models.Length);

            var rules = GetRules<T>();

            foreach (var model in models)
            {
                var row = sheet.CreateRow(newRowInd++);
                foreach (var rule in rules)
                {
                    var value = rule.Prop.GetValue(model);
                    if (value is string strValue)
                    {
                        row.CreateCell(rule.Attribute.ColumnIndex).SetCellValue(strValue);
                    }
                    else if (value is DateTime dateValue)
                    {
                        row.CreateCell(rule.Attribute.ColumnIndex).SetCellValue(dateValue);
                    }
                    else if (value is bool boolValue)
                    {
                        row.CreateCell(rule.Attribute.ColumnIndex).SetCellValue(boolValue);
                    }
                    else if (double.TryParse(value.ToString(), out double numValue))
                    {
                        row.CreateCell(rule.Attribute.ColumnIndex).SetCellValue(numValue);
                    }
                    else if (value is Enum enumValue)
                    {
                        FieldInfo fi = enumValue.GetType().GetField(enumValue.ToString());

                        var attributes = (DescriptionAttribute[])fi.GetCustomAttributes(
                            typeof(DescriptionAttribute), false);

                        if (attributes != null && attributes.Length > 0)
                            row.CreateCell(rule.Attribute.ColumnIndex).SetCellValue(attributes[0].Description);
                    }
                    else
                    {
                        row.CreateCell(rule.Attribute.ColumnIndex).SetCellValue(value.ToString());
                    }
                }
            }

            SaveExcel(workbook, toFile);
        }

        private void SaveExcel(XSSFWorkbook workbook, string filePath)
        {
            using (var file = new FileStream(filePath, FileMode.OpenOrCreate))
            {
                workbook.Write(file, true);
            }
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

        private BindProp<WriteColAttribute>[] GetRules<T>()
        {
            return typeof(T)
                .GetProperties()
                .Select(p => new BindProp<WriteColAttribute>(p, p.GetCustomAttribute(typeof(WriteColAttribute), true) as WriteColAttribute))
                .Where(a => a.Attribute != null)
                .ToArray();
        }
    }
}
