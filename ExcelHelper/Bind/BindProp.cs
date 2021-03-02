using ExcelHelper.Bind.Binders;

namespace ExcelHelper.Bind
{
    internal class BindProp
    {
        internal BindProp(string propName, BaseColAttribute attribute)
        {
            PropName = propName;
            Attribute = attribute;
        }

        internal string PropName { get; set; }
        internal BaseColAttribute Attribute { get; set; }
    }
}
