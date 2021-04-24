using System;
using System.ComponentModel;
using System.Linq.Expressions;
using System.Reflection;

namespace ExcelEnt.Extentions
{
    internal static class TypeExtentions
    {
        internal static string DescriptionAttrValue(this MemberInfo propertyInfo)
        {
            var attributes = (DescriptionAttribute[])propertyInfo.GetCustomAttributes(
                typeof(DescriptionAttribute), false);

            if (attributes != null && attributes.Length > 0)
                return attributes[0].Description;

            return null;
        }

        internal static PropertyInfo GetProperty<T, P>(Expression<Func<T, P>> property)
        {
            LambdaExpression lambda = (LambdaExpression)property;
            MemberExpression memberExpression;

            if (lambda.Body is UnaryExpression)
            {
                UnaryExpression unaryExpression = (UnaryExpression)(lambda.Body);
                memberExpression = (MemberExpression)(unaryExpression.Operand);
            }
            else
            {
                memberExpression = (MemberExpression)(lambda.Body);
            }

            return (PropertyInfo)memberExpression.Member;
        }
    }
}
