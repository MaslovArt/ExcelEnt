namespace ExcelHelper.Extentions
{
    internal static class EnumExtentions
    {
        internal static string ToDescription<T>(this T _enum) =>
            _enum.GetType().GetField(_enum.ToString()).DescriptionAttrValue();
    }
}
