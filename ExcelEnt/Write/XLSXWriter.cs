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
        private List<WriteRule>                     _rules;
        private XLSXStyling<T>                      _styling;
        private XLSXTemplating<T>                   _templating;
        private List<Action<XSSFWorkbook, ISheet>>  _modifications;

        public XLSXWriter()
        {
            _rules = new List<WriteRule>();
            _templating = new XLSXTemplating<T>();
            _styling = new XLSXStyling<T>();
            _modifications = new List<Action<XSSFWorkbook, ISheet>>();
        }

        public XLSXWriter<T> AddRule(Expression<Func<T, object>> propName, int colIndex)
        {
            var prop = TypeExtentions.GetProperty(propName);
            _rules.Add(new WriteRule(colIndex, prop));

            return this;
        }
        
        public XLSXWriter<T> UseTemplating(Action<XLSXTemplating<T>> config)
        {
            config(_templating);

            return this;
        }

        public XLSXWriter<T> UseStyling(Action<XLSXStyling<T>> config)
        {
            config(_styling);

            return this;
        }

        public XLSXWriter<T> Modify(Action<XSSFWorkbook, ISheet> modification)
        {
            _modifications.Add(modification);

            return this;
        }

        public void Generate(string resultFilePath, T[] entities)
        {
            var insertIndex = _templating?.InsertInd ?? 0;

            var workbook = _templating?.CreateWorkbook(entities.Length) ?? new XSSFWorkbook();
            var sheet = workbook.GetSheetAt(0);

            _styling?.Build(workbook);
            WriteEntities(sheet, entities, insertIndex);
            _styling?.ApplyConditionRowStyles(sheet, entities, insertIndex);
            ApplyModifications(workbook, sheet);

            SaveExcel(workbook, resultFilePath);
        }


        private void ApplyModifications(XSSFWorkbook workbook, ISheet sheet)
        {
            foreach (var modification in _modifications)
                modification(workbook, sheet);
        }

        private void WriteEntities(ISheet sheet, T[] entities, int insertIndex)
        {
            var newRowInd = insertIndex;
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
            _styling?.SetStyle(newCell);

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
