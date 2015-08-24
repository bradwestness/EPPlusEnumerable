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
    using System.Collections.Concurrent;
    using System.Data;
    using System.Threading.Tasks;

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
            var enumeratedData = data.ToList();

            foreach (var collection in enumeratedData)
            {
                AddWorksheet(package, collection.ToList());
            }

            AddSpreadsheetLinks(package, enumeratedData);

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
            var enumeratedData = data.ToList();

            foreach (var collection in enumeratedData)
            {
                AddWorksheet(package, collection.ToList());
            }

            AddSpreadsheetLinks(package, enumeratedData);

            return package;
        }

        /// <summary>
        /// Creates an Excel spreadsheet with a single worksheet for the supplied data.
        /// </summary>
        /// <param name="data">Each row of the spreadsheet will contain one item from the data collection.</param>
        /// <returns>An Excel spreadsheet as a byte array.</returns>
        public static byte[] Create(IEnumerable<object> data)
        {
            var package = new ExcelPackage();
            var enumeratedData = data.ToList();

            AddWorksheet(package, enumeratedData);
            AddSpreadsheetLinks(package, new[] { enumeratedData });

            return package.GetAsByteArray();
        }

        /// <summary>
        /// Creates an Excel spreadsheet with a single worksheet for the supplied data.
        /// </summary>
        /// <param name="data">Each row of the spreadsheet will contain one item from the data collection.</param>
        /// <returns>A populated ExcelPackage object.</returns>
        public static ExcelPackage CreatePackage(IEnumerable<object> data)
        {
            var package = new ExcelPackage();
            var enumeratedData = data.ToList();

            AddWorksheet(package, enumeratedData);
            AddSpreadsheetLinks(package, new[] { enumeratedData });

            return package;
        }

        #endregion

        #region Private Methods

        private static ExcelWorksheet AddWorksheet(ExcelPackage package, List<object> data)
        {
            if (data == null || !data.Any())
            {
                return null;
            }

            string skipProperty = null;
            var firstRow = data.First();
            var collectionType = firstRow.GetType();
            var properties = collectionType.GetProperties();
            var worksheetName = GetWorksheetName(firstRow, collectionType, out skipProperty);
            var worksheet = package.Workbook.Worksheets.Add(worksheetName);
            var col = 0;

            var visibleProperties = new List<PropertyInfo>();
            var dataTable = new DataTable();


            // add column headings
            for (var i = 0; i < properties.Count(); i++)
            {
                var property = properties[i];
                var propertyName = GetPropertyName(property);

                if (skipProperty != null && property.Name.Equals(skipProperty))
                {
                    continue;
                }

                dataTable.Columns.Add(propertyName);
                visibleProperties.Add(property);

                col += 1;
                worksheet.Cells[string.Format("{0}1", GetColumnLetter(col))].Value = propertyName;
            }

            var dataRows = new ConcurrentDictionary<long, List<object>>(); 

            Parallel.ForEach(
                data,
                (item, pls, index) =>
                {
                    List<object> dataRow = new List<object>();

                    foreach (var visibleProperty in visibleProperties)
                    {
                        object propertyValue = GetPropertyValue(visibleProperty, item);

                        dataRow.Add(propertyValue);
                    }

                    dataRows.GetOrAdd(index, dataRow);
                });


            // add rows to data table
            for (int i = 0; i < data.Count; i++)
            {
                dataTable.Rows.Add(dataRows[i].ToArray());
            }

            // use EPP LoadFromDataTable method in second row
            worksheet.Cells["A2"].LoadFromDataTable(dataTable, false);


            // set table formatting
            using (var range = worksheet.Cells[string.Format("A1:{0}{1}", GetColumnLetter(col), data.Count() + 1)])
            {
                range.AutoFitColumns();

                var table = worksheet.Tables.Add(range, "table_" + worksheetName);
                table.TableStyle = GetTableStyle(collectionType);
            }

            return worksheet;
        }

        private static string GetWorksheetName(object firstRow, Type collectionType, out string skipProperty)
        {
            skipProperty = null;
            var worksheetName = collectionType.Name;

            // this is just to strip out the giant string of numbers that EntityFramework appends to
            // proxy types if you passed in a collection of those rather than a type with a
            // displayname attribute specified
            if (worksheetName.Contains('_'))
            {
                worksheetName = worksheetName.Substring(0, worksheetName.IndexOf('_'));
            }

            var worksheetNameProperty = collectionType
                .GetProperties()
                .FirstOrDefault(prop => Attribute.IsDefined(prop, typeof(SpreadsheetTabNameAttribute), true));

            if (worksheetNameProperty != null)
            {
                var worksheetNameAttribute = worksheetNameProperty.GetCustomAttribute<SpreadsheetTabNameAttribute>(true);
                var worksheetPropertyValue = worksheetNameProperty.GetValue(firstRow);

                if (!string.IsNullOrWhiteSpace(worksheetNameAttribute.FormatString))
                {
                    worksheetName = string.Format(worksheetNameAttribute.FormatString, worksheetPropertyValue);
                }
                else
                {
                    worksheetName = worksheetPropertyValue.ToString();
                }

                if (worksheetNameAttribute.ExcludeFromOutput)
                {
                    // this property has a SpreadsheetTabNameAttribute
                    // with ExcludeFromOutput set to true
                    skipProperty = worksheetNameProperty.Name;
                }
            }
            else
            {
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
            var inputValue = property.GetValue(item);
            object outputValue = string.Empty;

            var displayFormatAttribute = property.GetCustomAttribute<DisplayFormatAttribute>(true);
            if (displayFormatAttribute != null)
            {
                if (inputValue == null && !string.IsNullOrWhiteSpace(displayFormatAttribute.NullDisplayText))
                {
                    outputValue = displayFormatAttribute.NullDisplayText;
                }
                else if (inputValue != null)
                {
                    outputValue = string.Format(displayFormatAttribute.DataFormatString, inputValue);
                }
            }
            else if (inputValue != null)
            {
                if (property.PropertyType.IsValueType)
                {
                    // for value types, just output the raw value
                    outputValue = inputValue;
                }
                else
                {
                    var enumerable = (inputValue as IEnumerable<object>);
                    if (enumerable != null)
                    {
                        // for collections, return a count
                        outputValue = enumerable.Count();
                    }
                    else
                    {
                        // reference type
                        outputValue = inputValue.ToString();
                    }
                }
            }

            return outputValue;
        }

        private static void AddSpreadsheetLinks(ExcelPackage package, IEnumerable<IEnumerable<object>> data)
        {
            foreach (var collection in data)
            {
                var enumeratedCollection = collection.ToList();

                if (!enumeratedCollection.Any())
                {
                    continue;
                }

                string skipProperty = null;
                var firstRow = enumeratedCollection.First();
                var collectionType = firstRow.GetType();
                var properties = collectionType.GetProperties();
                var worksheetName = GetWorksheetName(firstRow, collectionType, out skipProperty);
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

                    if (skipProperty != null && property.Name.Equals(skipProperty))
                    {
                        // this property has a SpreadsheetTabName attribute
                        // with ExcludeFromOutput set to true
                        continue;
                    }

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