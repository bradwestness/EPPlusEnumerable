using OfficeOpenXml;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Drawing;
using System.Linq;
using System.Reflection;

namespace EPPlusEnumerable
{
    public static class Spreadsheet
    {
        #region Static Fields

        private static readonly char[] _letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".ToArray();

        private const TableStyles DefaultTableStyle = TableStyles.Medium16;

        #endregion

        #region Public Methods

        /// <summary>
        /// Creates an Excel spreadsheet with worksheets for each collection of objects.
        /// </summary>
        /// <param name="data">A collection of data collections. Each outer collection will be used as a worksheet, while the inner collections will be used as data rows.</param>
        /// <returns>An Excel spreadsheet as a byte array.</returns>
        public static byte[] Create(IEnumerable<IEnumerable<object>> data)
        {
            var package = new ExcelPackage();

            foreach (var collection in data)
            {
                AddWorksheet(package, collection);
            }

            AddSpreadsheetLinks(package, data);

            return package.GetAsByteArray();
        }

        /// <summary>
        /// Creates an Excel spreadsheet with worksheets for each collection of objects using custom names for the worksheets.
        /// </summary>
        /// <param name="data">A collection of data collections. Each outer collection will be used as a worksheet using the IDictionary key as the sheet name, while the inner collections will be used as data rows.</param>
        /// <returns>An Excel spreadsheet as a byte array.</returns>
        public static byte[] Create(IDictionary<string,IEnumerable<object>> data)
        {
            var package = new ExcelPackage();

            foreach(var collection in data)
            {
                AddWorksheet(package, collection.Value, collection.Key);
            }

            AddSpreadsheetLinks(package, data.Values);

            return package.GetAsByteArray();
        }

        /// <summary>
        /// Creates an Excel spreadsheet with worksheets for each collection of objects.
        /// </summary>
        /// <param name="data">A collection of data collections. Each outer collection will be used as a worksheet, while the inner collections will be used as data rows.</param>
        /// <returns>A populated ExcelPackage object.</returns>
        public static ExcelPackage CreatePackage(IEnumerable<IEnumerable<object>> data)
        {
            var package = new ExcelPackage();

            foreach (var collection in data)
            {
                AddWorksheet(package, collection);
            }

            AddSpreadsheetLinks(package, data);

            return package;
        }

        /// <summary>
        /// Creates an Excel spreadsheet with worksheets for each collection of objects.
        /// </summary>
        /// <param name="data">A collection of data collections. Each outer collection will be used as a worksheet using the IDictionary key as the sheet name, while the inner collections will be used as data rows.</param>
        /// <returns>A populated ExcelPackage object.</returns>
        public static ExcelPackage CreatePackage(IDictionary<string,IEnumerable<object>> data)
        {
            var package = new ExcelPackage();

            foreach(var collection in data)
            {
                AddWorksheet(package, collection.Value, collection.Key);
            }

            AddSpreadsheetLinks(package, data.Values);

            return package;
        }

        /// <summary>
        /// Creates an Excel spreadsheet with a single worksheet for the supplied data.
        /// </summary>
        /// <param name="data">Each row of the spreadsheet will contain one item from the data collection.</param>
        /// <param name="worksheetName">The worksheet name to use. If omitted it will use the default naming conventions.</param>
        /// <returns>An Excel spreadsheet as a byte array.</returns>
        public static byte[] Create(IEnumerable<object> data, string worksheetName = null)
        {
            var package = new ExcelPackage();

            AddWorksheet(package, data, worksheetName);
            AddSpreadsheetLinks(package, new[] { data });

            return package.GetAsByteArray();
        }

        /// <summary>
        /// Creates an Excel spreadsheet with a single worksheet for the supplied data.
        /// </summary>
        /// <param name="data">Each row of the spreadsheet will contain one item from the data collection.</param>
        /// <param name="worksheetName">The worksheet name to use. If omitted it will use the default naming conventions.</param>
        /// <returns>A populated ExcelPackage object.</returns>
        public static ExcelPackage CreatePackage(IEnumerable<object> data, string worksheetName = null)
        {
            var package = new ExcelPackage();

            AddWorksheet(package, data, worksheetName);
            AddSpreadsheetLinks(package, new[] { data });

            return package;
        }

        #endregion

        #region Private Methods

        private static ExcelWorksheet AddWorksheet(ExcelPackage package, IEnumerable<object> data, string worksheetName = null)
        {
            if (data == null || !data.Any())
            {
                return null;
            }

            var collectionType = data.First().GetType();
            var properties = collectionType.GetProperties();
            var worksheet = package.Workbook.Worksheets.Add(worksheetName);
            var lastColumn = GetColumnLetter(properties.Count());

            //Get worksheet name if it has not been specified already
            if (string.IsNullOrWhiteSpace(worksheetName))
            {
                worksheetName = GetWorksheetName(collectionType);
            }

            // add column headings
            for (var i = 1; i <= properties.Count(); i++)
            {
                var property = properties[i - 1];
                var propertyName = GetPropertyName(property);

                worksheet.Cells[string.Format("{0}1", GetColumnLetter(i))].Value = propertyName;
            }

            // add rows (starting with two, since Excel is 1-based and we added a row of column headings)
            for (var row = 2; row < data.Count() + 2; row++)
            {
                var item = data.ElementAt(row - 2);

                for (var col = 1; col <= properties.Count(); col++)
                {
                    var cell = string.Format("{0}{1}", GetColumnLetter(col), row);
                    var property = properties.ElementAt(col - 1);
                    var value = property.GetValue(item) ?? string.Empty;

                    switch (property.PropertyType.Name.ToLower())
                    {
                        // TODO: potentially add special formatting for other types (images, etc)

                        case "icollection`1":
                            // if the property is another collection, just show the count
                            worksheet.Cells[cell].Value = (value as IEnumerable<object>).Count();
                            break;

                        default:
                            var propertyValue = GetPropertyValue(property, item);
                            worksheet.Cells[cell].Value = propertyValue;
                            if(propertyValue is DateTime){worksheet.Cells[cell].Style.Numberformat.Format = System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern;}
                            break;
                    }
                }
            }

            // set table formatting
            using (var range = worksheet.Cells[string.Format("A1:{0}{1}", lastColumn, data.Count() + 1)])
            {
                range.AutoFitColumns();

                var table = worksheet.Tables.Add(range, "table_" + worksheetName);
                table.TableStyle = GetTableStyle(collectionType);
            }

            return worksheet;
        }

        private static string GetWorksheetName(Type collectionType)
        {
            var worksheetName = collectionType.Name;

            // this is just to strip out the giant string of numbers that EntityFramework appends to
            // proxy types if you passed in a collection of those rather than a type with a
            // displayname attribute specified
            if (worksheetName.Contains('_'))
            {
                worksheetName = worksheetName.Substring(0, worksheetName.IndexOf('_'));
            }

            var displayNameAttribute = collectionType.GetCustomAttribute<DisplayNameAttribute>(true);
            if (displayNameAttribute != null)
            {
                worksheetName = displayNameAttribute.DisplayName;
            }
            else
            {
                var displayAttribute = collectionType.GetCustomAttribute<DisplayAttribute>(true);
                if (displayAttribute != null)
                {
                    worksheetName = displayAttribute.Name;
                }
            }

            return worksheetName;
        }

        private static TableStyles GetTableStyle(Type collectionType)
        {
            var tableStyle = DefaultTableStyle;

            var spreadsheetTableStyleAttribute = collectionType.GetCustomAttribute<SpreadsheetTableStyleAttribute>(true);
            if (spreadsheetTableStyleAttribute != null)
            {
                tableStyle = spreadsheetTableStyleAttribute.TableStyle;
            }

            return tableStyle;
        }

        private static string GetPropertyName(PropertyInfo property)
        {
            var propertyName = property.Name;

            var displayNameAttribute = property.GetCustomAttribute<DisplayNameAttribute>(true);
            if (displayNameAttribute != null)
            {
                propertyName = displayNameAttribute.DisplayName;
            }
            else
            {
                var displayAttribute = property.GetCustomAttribute<DisplayAttribute>(true);
                if (displayAttribute != null)
                {
                    propertyName = displayAttribute.Name;
                }
            }

            return propertyName;
        }

        private static object GetPropertyValue(PropertyInfo property, object item)
        {
            var value = property.GetValue(item);
            object returnValue = string.Empty;

            var displayFormatAttribute = property.GetCustomAttribute<DisplayFormatAttribute>(true);
            if (displayFormatAttribute != null)
            {
                if (value == null && !string.IsNullOrWhiteSpace(displayFormatAttribute.NullDisplayText))
                {
                    returnValue = displayFormatAttribute.NullDisplayText;
                }
                else if (value != null)
                {
                    returnValue = string.Format(displayFormatAttribute.DataFormatString, value);
                }
            }
            else if (value != null)
            {
                //If it's a simple type, just pass it along. Otherwise, send back the string representation
                switch(Type.GetTypeCode(value.GetType()))
                {
                    case TypeCode.Boolean:
                    case TypeCode.Int16:
                    case TypeCode.UInt16:
                    case TypeCode.Int32:
                    case TypeCode.UInt32:
                    case TypeCode.Int64:
                    case TypeCode.UInt64:
                    case TypeCode.Single:
                    case TypeCode.Double:
                    case TypeCode.Decimal:
                    case TypeCode.DateTime:
                    case TypeCode.String:
                    returnValue = value;
                        break;
                    default:
                        returnValue = value.ToString();
                        break;
                }
            }

            return returnValue;
        }

        private static void AddSpreadsheetLinks(ExcelPackage package, IEnumerable<IEnumerable<object>> data)
        {
            foreach (var collection in data)
            {
                if (collection == null || !collection.Any())
                {
                    continue;
                }

                var collectionType = collection.First().GetType();
                var properties = collectionType.GetProperties();
                var worksheetName = GetWorksheetName(collectionType);
                var worksheet = package.Workbook.Worksheets[worksheetName];

                if (worksheet == null)
                {
                    continue;
                }

                // loop through the properties in the collection type
                // and see if any have a SpreadsheetLinkAttribute specified
                for (var prop = 1; prop <= properties.Count(); prop++)
                {
                    var property = properties.ElementAt(prop - 1);
                    var attribute = property.GetCustomAttribute<SpreadsheetLinkAttribute>();

                    if (attribute == null)
                    {
                        // no SpreadsheetLinkAttribute for this property,
                        // so skip to the next property
                        continue;
                    }

                    // get the worksheet specified by the attribute
                    var linkSheet = package.Workbook.Worksheets[attribute.WorksheetName];

                    if (linkSheet == null)
                    {
                        // if the target worksheet specified by the attribute doesn't exist,
                        // we can't add any links, so just skip to the next property
                        continue;
                    }

                    var linkColumn = string.Empty;

                    // loop through the columns of the first row of the
                    // attribute's target worksheet and see if any
                    // match the attribute's column header value
                    for (var col = 1; col <= linkSheet.Dimension.Columns; col++)
                    {
                        var letter = GetColumnLetter(col);
                        if (linkSheet.Cells[letter + "1"].Value.ToString().Equals(attribute.ColumnHeaderValue))
                        {
                            linkColumn = letter;
                            break;
                        }
                    }

                    if (!string.IsNullOrWhiteSpace(linkColumn))
                    {
                        // we found the target column of the target worksheet, so we can add links!
                        AddColumnLinks(worksheet, prop, linkSheet, linkColumn);
                    }
                }
            }
        }

        private static void AddColumnLinks(ExcelWorksheet worksheet, int worksheetColumnIndex, ExcelWorksheet linkSheet, string linkColumn)
        {
            // loop through the cells of the worksheet that correspond to the property
            // that had the SpreadsheetLinkAttribute and try to find a link target for each
            for (var worksheetRow = 1; worksheetRow <= worksheet.Dimension.Rows; worksheetRow++)
            {
                var worksheetCell = worksheet.Cells[string.Format("{0}{1}", GetColumnLetter(worksheetColumnIndex), worksheetRow)];
                var worksheetValue = worksheetCell.Value.ToString();

                // loop through the cells of the target worksheet column and see if any of the values
                // match the value of the current worksheet cell
                for (var linksheetRow = 1; linksheetRow <= linkSheet.Dimension.Rows; linksheetRow++)
                {
                    var linksheetValue = linkSheet.Cells[string.Format("{0}{1}", linkColumn, linksheetRow)].Value.ToString();

                    if (worksheetValue.Equals(linksheetValue))
                    {
                        // we found a match! this is the link target, 
                        // so add the hyperlink to the worksheet cell
                        // and stop searching for targets for this row
                        worksheetCell.Hyperlink = new ExcelHyperLink(string.Format("'{0}'!{1}{2}", linkSheet.Name, linkColumn, linksheetRow), worksheetValue.ToString());
                        worksheetCell.Style.Font.UnderLine = true;
                        worksheetCell.Style.Font.Color.SetColor(Color.Blue);
                        break;
                    }
                }
            }
        }

        /// <summary>
        /// Gets the Excel-style column letter for the specified numerical index (e.g. 4 is D, 26 is Z, 27 is AA, 28 is AB...).
        /// </summary>
        /// <param name="column">The numerical index of the column.</param>
        /// <returns>The corresponding Excel-style column letter.</returns>
        private static string GetColumnLetter(int column)
        {
            if (column <= _letters.Length)
            {
                return _letters[column - 1].ToString();
            }

            var number = column;
            string letter = string.Empty;

            while (number > 0)
            {
                var remainder = (number - 1) % _letters.Length;
                letter = _letters[remainder] + letter;
                number = (number - remainder) / _letters.Length;
            }

            return letter;
        }

        #endregion
    }
}
