using System;

namespace EPPlusEnumerable
{
    /// <summary>
    /// Use this attribute to denote that the property is to be excluded and not outputted in the Excel worksheet.
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
    public class SpreadsheetExcludeAttribute : Attribute
    {
        #region Constructors

        public SpreadsheetExcludeAttribute()
        {

        }

        #endregion
    }
}
