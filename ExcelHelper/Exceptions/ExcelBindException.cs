using System;

namespace ExcelHelper.Exceptions
{
    internal class ExcelBindException : Exception
    {
        public ExcelBindException(int row, int col, Exception ex)
        {
            Row = row;
            Col = col;
            Inner = ex;
        }

        public int Row { get; set; }
        public int Col { get; set; }
        public Exception Inner { get; set; }
    }
}
