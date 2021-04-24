using System;

namespace ExcelEnt.Exceptions
{
    internal class ExcelBindException : Exception
    {
        public ExcelBindException(int row, int col, Exception ex)
            :base(ex.Message, ex)
        {
            Row = row;
            Col = col;
        }

        public int Row { get; set; }
        public int Col { get; set; }
    }
}
