using ExcelHelper.Bind.Binders;
using NPOI.XSSF.UserModel;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

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

            var mapRules = GetRules<T>();

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
                    var cellValue = rowItem.GetCell(rule.Attribute.ColumnIndex);
                    rule.Attribute.MapProp(rule.PropName, cellValue, newModel);
                }
                models.Add(newModel);
            }

            return models.ToArray();
        }

        private BindProp[] GetRules<T>()
        {
            return typeof(T)
                .GetProperties()
                .Select(p => new BindProp(p.Name, p.GetCustomAttribute(typeof(BaseColAttribute), true) as BaseColAttribute))
                .Where(a => a.Attribute != null)
                .ToArray();
        }
    }
}
