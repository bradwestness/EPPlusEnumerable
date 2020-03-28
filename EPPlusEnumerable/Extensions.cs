using System;
using System.Reflection;

namespace EPPlusEnumerable
{
    internal static class Extensions
    {
        public static T GetCustomAttribute<T>(this MemberInfo element, bool inherit = true) where T : Attribute =>
            Attribute.GetCustomAttribute(element, typeof(T), inherit) as T;

        public static object GetValue(this PropertyInfo element, object obj) =>
            element.GetValue(obj, BindingFlags.Default, null, null, null);
    }
}
