using System.Linq.Expressions;
using System.Reflection;

namespace LazySpreadsheets.Extensions;

internal static class ExpressionExtensions
{
    /// <summary>
    /// Gets the property info of the property on the declaring type.
    /// </summary>
    /// <param name="propertyExpression">Expression that selects a property</param>
    /// <returns>The selected property info</returns>
    internal static PropertyInfo GetPropertyInfo(Expression propertyExpression)
    {
        if (propertyExpression.NodeType != ExpressionType.Lambda)
        {
            throw new ArgumentException("Selector must be lambda expression", nameof(propertyExpression));
        }

        var lambda = (LambdaExpression)propertyExpression;

        var memberExpression = ExtractMemberExpression(lambda.Body);

        if (memberExpression == null)
        {
            throw new ArgumentException("Selector must be member access expression", nameof(propertyExpression));
        }

        if (memberExpression.Member.DeclaringType == null)
        {
            throw new InvalidOperationException("Property does not have declaring type");
        }

        return memberExpression.Member.DeclaringType.GetProperty(memberExpression.Member.Name)!;
    }

    private static MemberExpression? ExtractMemberExpression(Expression expression)
    {
        switch (expression.NodeType)
        {
            case ExpressionType.MemberAccess:
                return ((MemberExpression)expression);
            case ExpressionType.Convert:
                var operand = ((UnaryExpression)expression).Operand;
                return ExtractMemberExpression(operand);
            default:
                return null;
        }
    }
}