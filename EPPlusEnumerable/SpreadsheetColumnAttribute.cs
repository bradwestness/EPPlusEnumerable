using System;

namespace EPPlusEnumerable
{
    /// <summary>
    /// Use this attribute to explicitly denote that the property is a spreadsheet column.
    /// A format string may optionally be supplied, and setting ExcludeFromOutput to true will
    /// cause the resulting spreadsheet to not include a column for this property.
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
    public class SpreadsheetColumnAttribute : Attribute
    {
        #region Properties

        public bool ExcludeFromOutput { get; set; }

        public string FormatString { get; set; }

        #endregion

        #region Constructors

        public SpreadsheetColumnAttribute()
        {
            ExcludeFromOutput = false;
            FormatString = null;
        }

        #endregion
    }
}
