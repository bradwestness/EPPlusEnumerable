namespace EPPlusEnumerable
{
    using System;

    /// <summary>
    /// Use this attribute to denote that the property is not used to create a column in an excel worksheet when exporting data.
    /// </summary>
    [AttributeUsage(AttributeTargets.Property,AllowMultiple = false,Inherited = true)]
    public class SpreadsheetExcludeFromOutputAttribute : Attribute
    {   
    }
}