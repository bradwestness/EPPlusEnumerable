namespace EPPlusEnumerable
{
    using System;
    using System.Reflection;

    internal static class Extensions
    {
        public static T GetCustomAttribute<T>(this MemberInfo element, Type type, bool inherit) where T : Attribute
        {
            var attr = (T)Attribute.GetCustomAttribute(element, typeof(T), inherit);
            
            if (attr == null)
            {
                return element.GetMetaAttribute<T>(type, inherit);
            }
            
            return attr;
        }

        public static T GetCustomAttribute<T>(this MemberInfo element, Type type) where T : Attribute
        {
            var attr = (T)Attribute.GetCustomAttribute(element, typeof(T));
            
            if (attr == null)
            {
                return element.GetMetaAttribute<T>(type);
            }
            
            return attr;
        }

        public static object GetValue(this PropertyInfo element, Object obj)
        {
            return element.GetValue(obj, BindingFlags.Default, null, null, null);
        }
    }
}
