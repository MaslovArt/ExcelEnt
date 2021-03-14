using ExcelEnt.Extentions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Linq.Expressions;

namespace ExcelEnt.Write
{
    /// <summary>
    /// Entities to excel writer
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public class XLSXWriter<T>
    {
        internal List<WriteRule>    _rules;
        internal XLSXStyling<T>     _styling;
        internal XLSXTemplating<T>  _templating;

        public XLSXWriter()
        {
            _rules = new List<WriteRule>();
            _templating = new XLSXTemplating<T>();
            //_styling = new XLSXStyling<T>(_workbook, _sheet);
        }

        private int InsertIndex => _templating.InsertInd;

        /// <summary>
        /// Add entities to excel value rule
        /// </summary>
        /// <param name="propName">Entity property name</param>
        /// <param name="colIndex">Excel cell index</param>
        /// <param name="styleName">Existing style name</param>
        /// <returns></returns>
        public XLSXWriter<T> AddRule(Expression<Func<T, object>> propName, int colIndex, string styleName = null)
        {
            var prop = TypeExtentions.GetProperty(propName);
            _rules.Add(new WriteRule(colIndex, prop, styleName));

            return this;
        }
        public XLSXWriter<T> UseTemplating(Action<XLSXTemplating<T>> config)
        {
            _templating = new XLSXTemplating<T>();
            config(_templating);

            return this;
        }

        public XLSXWriter<T> UseStyling(Action<XLSXStyling<T>> config)
        {
            config(_styling);

            return this;
        }

        /// <summary>
        /// Workbook and sheet modifications
        /// </summary>
        /// <param name="action">Custom modification</param>
        /// <returns></returns>
        public XLSXWriter<T> Modify(Action<XSSFWorkbook, ISheet> action)
        {
            //action(_workbook, _sheet);

            return this;
        }

        /// <summary>
        /// Create excel file from entities
        /// </summary>
        /// <param name="resultFilePath">Created excel file path</param>
        /// <param name="entities">Entities</param>
        public void Generate(string resultFilePath, T[] entities)
        {
            var workbook = _templating?.CreateWorkbook(entities.Length) ?? new XSSFWorkbook();
            var sheet = workbook.GetSheetAt(0);

            WriteEntities(sheet, entities);

            //_styling.ApplyConditionRowStyles(entities, InsertIndex);

            SaveExcel(workbook, resultFilePath);
        }

        private void WriteEntities(ISheet sheet, T[] entities)
        {
            var newRowInd = InsertIndex;
            var minColIndex = _rules.Select(r => r.ExcelColInd).Min();
            var maxColIndex = _rules.Select(r => r.ExcelColInd).Max();

            foreach (var model in entities)
            {
                var row = sheet.CreateRow(newRowInd++);
                for (var colInd = minColIndex; colInd <= maxColIndex; colInd++)
                    CreateStyledCell(row, colInd);

                foreach (var rule in _rules)
                {
                    var value = rule.Prop.GetValue(model);
                    var newCell = row.GetCell(rule.ExcelColInd);

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
        }

        private ICell CreateStyledCell(IRow row, int cellIndex)
        {
            var newCell = row.CreateCell(cellIndex);
            //_styling.SetStyle(newCell, null);

            return newCell;
        }

        private void SaveExcel(XSSFWorkbook workbook, string resultFilePath)
        {
            using (var file = new FileStream(resultFilePath, FileMode.CreateNew))
            {
                workbook.Write(file, true);
            }
        }
    }
}
