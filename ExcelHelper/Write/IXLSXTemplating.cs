namespace ExcelHelper.Write
{
    public interface IXLSXTemplating<T> : IXLSXWriter<T>
    {
        IXLSXTemplating<T> ReplaceShortCode(string shortCode, string value);
    }
}
