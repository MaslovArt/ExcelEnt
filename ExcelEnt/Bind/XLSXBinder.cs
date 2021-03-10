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
        private int? StartIndex { get; set; }
        private int? EndIndex { get; set; }

        /// <summary>
        /// Add starting read index
        /// </summary>
        /// <param name="ind">Start index</param>
        /// <returns></returns>
        public XLSXBinder<T> StartFrom(int ind)
        {
            StartIndex = ind;
            return this;
        }

        /// <summary>
        /// Add ending read index
        /// </summary>
        /// <param name="ind">End index</param>
        /// <returns></returns>
        public XLSXBinder<T> EndOn(int ind)
        {
            EndIndex = ind;
            return this;
        }

        /// <summary>
        /// Add excel cell value to entity property rule
        /// </summary>
        /// <param name="colIndex">Excel cell index</param>
        /// <param name="propName">Entity property name</param>
        /// <param name="map">Cell value to property mapper</param>
        /// <returns></returns>
        public XLSXBinder<T> AddRule(int colIndex, Expression<Func<T, object>> propName, Func<ICell, object> map)
        {
            var prop = TypeExtentions.GetProperty(propName);
            Rules.Add(new BindRule(colIndex, prop, map));

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
