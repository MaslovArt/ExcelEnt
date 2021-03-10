namespace ExcelHelper.Write
{
    public interface IXLSXTemplating<T> : IXLSXWriter<T>
    {
        /// <summary>
        /// Replace template shortcodes
        /// </summary>
        /// <param name="shortCode">Shortcode</param>
        /// <param name="value">Replace value</param>
        /// <returns></returns>
        IXLSXTemplating<T> ReplaceShortCode(string shortCode, string value);
    }
}
