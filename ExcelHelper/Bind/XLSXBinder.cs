using ExcelHelper.Bind.Binders;
using ExcelHelper.Exceptions;
using ExcelHelper.Extentions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;

namespace ExcelHelper.Bind
{
    public class XLSXBinder
    {
        public T[] Bind<T>(string filePath, string sheet, int? start = null, int? end = null) where T : new()
        {
            var file = new FileInfo(filePath);
            return Bind<T>(file, sheet, start, end);
        }

        public T[] Bind<T>(FileInfo file, string sheet, int? start = null, int? end = null) where T : new()
        {
            if (!file.Exists) throw new FileNotFoundException(file.FullName);

            var mapRules = typeof(T).BindPropsAttrs<BaseColAttribute>();

            var excelPackage = new XSSFWorkbook(file);
            var models = new List<T>();
            var itemSheet = excelPackage.GetSheet(sheet);

            var fromRow = start ?? 0;
            var toRow = end ?? itemSheet.LastRowNum;

            for (int row = fromRow; row <= toRow; row++)
            {
                var newModel = new T();
                foreach (var rule in mapRules)
                {
                    try
                    {
                        var rowItem = itemSheet.GetRow(row);
                        var cellValue = rowItem.GetCell(rule.Attribute.ColumnIndex, MissingCellPolicy.RETURN_BLANK_AS_NULL);
                        rule.Attribute.MapProp(rule, cellValue, newModel);
                    }
                    catch (Exception ex)
                    {
                        throw new ExcelBindException(row, rule.Attribute.ColumnIndex, ex);
                    }
                }
                models.Add(newModel);
            }

            return models.ToArray();
        }
    }
}
