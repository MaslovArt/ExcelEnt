using ExcelHelper.Exceptions;
using ExcelHelper.Extentions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq.Expressions;

namespace ExcelHelper.Bind
{
    public class XLSXBinder<T> where T : new()
    {
        private List<BindRule> Rules = new List<BindRule>();
        private int? StartIndex { get; set; }
        private int? EndIndex { get; set; }

        public XLSXBinder<T> StartFrom(int ind)
        {
            StartIndex = ind;
            return this;
        }

        public XLSXBinder<T> EndOn(int ind)
        {
            EndIndex = ind;
            return this;
        }

        public XLSXBinder<T> AddRule(int colIndex, Expression<Func<T, object>> propName, Func<ICell, object> map)
        {
            var prop = TypeExtentions.GetProperty(propName);
            Rules.Add(new BindRule(colIndex, prop, map));

            return this;
        } 

        public T[] Bind(string filePath, int pageInd)
        {
            var file = new FileInfo(filePath);
            return Bind(file, pageInd);
        }

        public T[] Bind(FileInfo file, int pageInd)
        {
            CheckFile(file);

            var excelPackage = new XSSFWorkbook(file);
            var itemSheet = excelPackage.GetSheetAt(pageInd);

            var fromRow = StartIndex ?? 0;
            var toRow = EndIndex ?? itemSheet.LastRowNum;

            var models = new List<T>();

            for (int rowInd = fromRow; rowInd <= toRow; rowInd++)
            {
                var newModel = new T();
                foreach (var rule in Rules)
                {
                    try
                    {
                        var row = itemSheet.GetRow(rowInd);
                        var cell = row.GetCell(rule.ExcelColInd);
                        var mappedValue = rule.Map(cell);

                        rule.Prop.SetValue(newModel, mappedValue);
                    }
                    catch (Exception ex)
                    {
                        throw new ExcelBindException(rowInd, rule.ExcelColInd, ex);
                    }
                }
                models.Add(newModel);
            }

            return models.ToArray();
        }

        private void CheckFile(FileInfo file)
        {
            if (!file.Exists) 
                throw new FileNotFoundException(file.FullName);
        }
    }
}
