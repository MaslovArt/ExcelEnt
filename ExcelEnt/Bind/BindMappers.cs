using ExcelEnt.Extentions;
using NPOI.SS.UserModel;
using System;

namespace ExcelEnt.Bind
{
    /// <summary>
    /// Excel cell value to property value mappers
    /// </summary>
    public static class BindMappers
    {
        /// <summary>
        /// Get int value
        /// </summary>
        /// <param name="cell">Excell cell</param>
        /// <returns></returns>
        public static int Int(ICell cell)
        {
            if (cell.CellType == CellType.Numeric)
                return (int)cell.NumericCellValue;

            return int.Parse(cell.ToString());
        }

        /// <summary>
        /// Get nullable int value
        /// </summary>
        /// <param name="cell">Excell cell</param>
        /// <returns></returns>
        public static int? NullInt(ICell cell) =>
            cell == null || string.IsNullOrEmpty(cell.ToString()) 
                ? null 
                : (int?)Int(cell);

        /// <summary>
        /// Get double value
        /// </summary>
        /// <param name="cell">Excell cell</param>
        /// <returns></returns>
        public static double Double(ICell cell)
        {
            if (cell.CellType == CellType.Numeric)
                return cell.NumericCellValue;

            return double.Parse(cell.ToString());
        }

        /// <summary>
        /// Get nullable double value
        /// </summary>
        /// <param name="cell">Excell cell</param>
        /// <returns></returns>
        public static double? NullDouble(ICell cell) =>
            cell == null || string.IsNullOrEmpty(cell.ToString()) 
                ? null 
                : (double?)Double(cell);

        /// <summary>
        /// Get string value
        /// </summary>
        /// <param name="cell">Excell cell</param>
        /// <returns></returns>
        public static string String(ICell cell) => 
            cell.ToString();

        /// <summary>
        /// Get date value
        /// </summary>
        /// <param name="cell">Excell cell</param>
        /// <returns></returns>
        public static DateTime Date(ICell cell) => 
            cell.DateCellValue;

        /// <summary>
        /// Get nullable date value
        /// </summary>
        /// <param name="cell">Excell cell</param>
        /// <returns></returns>
        public static DateTime? NullDate(ICell cell) =>
            cell == null || string.IsNullOrEmpty(cell.ToString()) 
                ? null 
                : (DateTime?)cell.DateCellValue;

        /// <summary>
        /// Get bool value
        /// </summary>
        /// <param name="cell">Excell cell</param>
        /// <returns></returns>
        public static bool Bool(ICell cell) => 
            cell.BooleanCellValue;

        /// <summary>
        /// Get nullable bool value
        /// </summary>
        /// <param name="cell">Excell cell</param>
        /// <returns></returns>
        public static bool? NullBool(ICell cell) =>
            string.IsNullOrEmpty(cell.ToString()) 
                ? null 
                : (bool?)cell.BooleanCellValue;

        /// <summary>
        /// Get bool value by true string value
        /// </summary>
        /// <param name="cell">Excell cell</param>
        /// <param name="trueValue">Excel cell value for true</param>
        /// <returns></returns>
        public static bool StringBool(ICell cell, string trueValue) => 
            cell.ToString() == trueValue;

        /// <summary>
        /// Get enum value
        /// </summary>
        /// <param name="cell">Excell cell</param>
        /// <returns></returns>
        public static Enum Enum<Enum>(ICell cell) => 
            cell.ToString().ToEnum<Enum>();
    }
}
