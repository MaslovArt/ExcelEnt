using ExcelEnt.Exceptions;
using ExcelEnt.Extentions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq.Expressions;

namespace ExcelEnt.Bind
{
    /// <summary>
    /// Excel sheet to entity binder
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public class XLSXBinder<T> where T : new()
    {
        private List<BindRule> Rules = new List<BindRule>();
        private int? SkipCount { get; set; }
        private int? TakeCount { get; set; }

        /// <summary>
        /// Add starting read index
        /// </summary>
        /// <param name="count">Start index</param>
        /// <returns></returns>
        public XLSXBinder<T> Skip(int count)
        {
            CheckCount(count, 0);
            SkipCount = count;

            return this;
        }

        /// <summary>
        /// Add ending read index
        /// </summary>
        /// <param name="count">End index</param>
        /// <returns></returns>
        public XLSXBinder<T> Take(int count)
        {
            CheckCount(count, 1);
            TakeCount = count;

            return this;
        }

        /// <summary>
        /// Add excel cell value to entity property rule
        /// </summary>
        /// <param name="colIndex">Excel cell index</param>
        /// <param name="propName">Entity property name</param>
        /// <param name="map">Cell value to property mapper</param>
        /// <returns></returns>
        public XLSXBinder<T> AddRule<P>(int colIndex, Expression<Func<T, P>> propName, Func<ICell, P> map)
        {
            var prop = TypeExtentions.GetProperty(propName);
            Rules.Add(new BindRule(colIndex, prop, (c) => map(c)));

            return this;
        } 

        /// <summary>
        /// Bind excel sheet to entity
        /// </summary>
        /// <param name="filePath">Excel file path</param>
        /// <param name="pageInd">Excel page index</param>
        /// <returns>Binded entities</returns>
        public T[] Bind(string filePath, int pageInd)
        {
            var file = new FileInfo(filePath);

            return Bind(file, pageInd);
        }

        /// <summary>
        /// Bind excel sheet to entity
        /// </summary>
        /// <param name="file">Excel file</param>
        /// <param name="pageInd">Excel page index</param>
        /// <returns>Binded entities</returns>
        public T[] Bind(FileInfo file, int pageInd)
        {
            CheckFile(file);

            var excelPackage = new XSSFWorkbook(file);
            var itemSheet = excelPackage.GetSheetAt(pageInd);

            var fromRow = SkipCount ?? 0;
            var toRow = SkipCount + TakeCount ?? itemSheet.LastRowNum + 1;

            var models = new List<T>();

            for (int rowInd = fromRow; rowInd < toRow; rowInd++)
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

        private void CheckCount(int count, int min)
        {
            if (count < min)
                throw new ArgumentException($"Count less then ${min}. (${count})");
        }

        private void CheckFile(FileInfo file)
        {
            if (!file.Exists) 
                throw new FileNotFoundException(file.FullName);
        }
    }
}
