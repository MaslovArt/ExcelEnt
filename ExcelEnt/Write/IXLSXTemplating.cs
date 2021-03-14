namespace ExcelEnt.Write
{
    public interface IXLSXTemplating<T>
    {
        IXLSXTemplating<T> FromTemplate(string filePath, int insertInd, bool moveFooter);
        IXLSXTemplating<T> ReplaceShortCode(string shortCode, string value);
    }
}