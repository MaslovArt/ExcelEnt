using ExcelHelper.Extentions;
using NPOI.SS.UserModel;
using System;

namespace ExcelHelper.Bind
{
    public static class BindMappers
    {
        public static object Int(ICell cell) => 
            (int)cell.NumericCellValue;

        public static object NullInt(ICell cell) =>
            cell == null ? null : (int?)cell.NumericCellValue;

        public static object Double(ICell cell) => 
            cell.NumericCellValue;

        public static object NullDouble(ICell cell) =>
            cell == null ? null : (double?)cell.NumericCellValue;

        public static string String(ICell cell) => 
            cell.ToString();

        public static object Date(ICell cell) => 
            cell.DateCellValue;

        public static object NullDate(ICell cell) =>
            cell == null ? null : (DateTime?)cell.DateCellValue;

        public static object Bool(ICell cell) => 
            cell.BooleanCellValue;

        public static object NullBool(ICell cell) =>
            cell == null ? null : (bool?)cell.BooleanCellValue;

        public static object StringBool(ICell cell, string trueValue) => 
            cell.ToString() == trueValue;

        public static object Enum<Enum>(ICell cell) => 
            cell.ToString().ToEnum<Enum>();

        public static object NullEnum<Enum>(ICell cell) where Enum : struct =>
            cell == null ? null : (Enum?)cell.ToString().ToEnum<Enum>();
    }
}
