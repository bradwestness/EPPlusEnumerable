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
        
        /// <summary>
        /// Crawls up the parent class to find any MetadaType to search for attributes as well.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="element">The property we're searchign for attributes on</param>
        /// <param name="type">The parent type of the property</param>
        /// <param name="inherit"></param>
        /// <returns></returns>
        public static T GetMetaAttribute<T>(this MemberInfo element, Type type, bool inherit = true) where T : Attribute
        {
            T metaAttr = null;
            // Get the MetadataType Attribute from the parent class
            var metaTypeAttrs = (MetadataTypeAttribute)type
                .GetCustomAttributes(typeof(MetadataTypeAttribute), inherit)
                .OfType<MetadataTypeAttribute>()
                .FirstOrDefault();

            // Check for null to avoid too much reflection
            if(metaTypeAttrs != null)
            {
                // Get the MetadataClass to check for attributes on
                var metaType = metaTypeAttrs.MetadataClassType;
                // Find the current element in the MetadataClass
                var metaElement = metaType?.GetMember(element.Name).FirstOrDefault();
                
                metaAttr = metaElement?.GetCustomAttribute<T>(inherit)
            }
            
            return metaAttr;
        }
    }
}
