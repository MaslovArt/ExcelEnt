using ExcelFileBinder.Binders;
using NPOI.XSSF.UserModel;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

namespace ExcelHelper
{
    public class XLSXHelper
    {
        public XLSXHelper()
        {
            var t = new XSSFWorkbook();
        }

        public T[] Bind<T>(FileInfo file, string sheet, int? start = null, int? end = null) where T : new()
        {
            if (!file.Exists) throw new FileNotFoundException(file.FullName);

            var mapRules = typeof(T)
                .GetProperties()
                .Select(p => new
                {
                    Prop = p.Name,
                    Attr = p.GetCustomAttribute(typeof(BaseColAttribute), true) as BaseColAttribute
                })
                .Where(a => a.Attr != null);

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
                    var rowItem = itemSheet.GetRow(row);
                    var cellValue = rowItem.GetCell(rule.Attr.ColumnIndex);
                    rule.Attr.MapProp(rule.Prop, cellValue, newModel);
                }
                models.Add(newModel);
            }

            return models.ToArray();
        }
    }
}
