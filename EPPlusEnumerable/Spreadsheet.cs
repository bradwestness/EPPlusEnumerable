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
        /// Creates an Excel spreadsheet with a single worksheet for the supplied data.
        /// </summary>
        /// <param name="data">Each row of the spreadsheet will contain one item from the data collection.</param>
        /// <returns>An Excel spreadsheet as a byte array.</returns>
        public static byte[] Create(IEnumerable<object> data)
        {
            var package = new ExcelPackage();

            AddWorksheet(package, data);
            AddSpreadsheetLinks(package, new[] { data });

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

            AddWorksheet(package, data);
            AddSpreadsheetLinks(package, new[] { data });

            return package;
        }

        #endregion

        #region Private Methods

        private static ExcelWorksheet AddWorksheet(ExcelPackage package, IEnumerable<object> data)
        {
            if (!(data is object) || !data.Any())
            {
                return null;
            }

            var firstRow = data.First();
            var collectionType = firstRow.GetType();
            var properties = collectionType.GetProperties();
            var worksheetName = GetWorksheetName(firstRow, collectionType);
            var worksheet = package.Workbook.Worksheets.Add(worksheetName);
            var col = 0;

            // add column headings
            for (var i = 0; i < properties.Count(); i++)
            {
                var property = properties[i];
                var propertyName = GetPropertyName(property, collectionType);

                if (property.GetCustomAttribute<SpreadsheetExcludeAttribute>() is object)
                {
                    // this property has a SpreadsheetExcludeAttribute
                    continue;
                }

                col += 1;
                worksheet.Cells[$"{GetColumnLetter(col)}1"].Value = propertyName;
            }

            // add rows (starting with two, since Excel is 1-based and we added a row of column headings)
            for (var row = 2; row < data.Count() + 2; row++)
            {
                var item = data.ElementAt(row - 2);
                col = 0;

                for (var i = 0; i < properties.Count(); i++)
                {
                    var property = properties.ElementAt(i);

                    if (property.GetCustomAttribute<SpreadsheetExcludeAttribute>() is object)
                    {
                        continue;
                    }

                    col += 1;
                    var cell = $"{GetColumnLetter(col)}{row}";

                    var cellStyleAttribute = property.GetCustomAttribute<SpreadsheetCellStyleAttribute>();
                    if (cellStyleAttribute is object)
                    {
                        cellStyleAttribute.ApplyCellStyles(worksheet.Cells[cell]);
                    }

                    _ = property.GetValue(item) ?? string.Empty;
                    worksheet.Cells[cell].Value = GetPropertyValue(property, item);
                }
            }

            // set table formatting
            using (var range = worksheet.Cells[$"A1:{GetColumnLetter(col)}{data.Count() + 1}"])
            {
                range.AutoFitColumns();

                var tableName = GetTableName(worksheetName);
                var table = worksheet.Tables.Add(range, $"table_{tableName}");
                table.TableStyle = GetTableStyle(collectionType);
            }

            return worksheet;
        }

        // Sanitize table name to only valid characters
        // https://support.office.com/en-us/article/Rename-an-Excel-table-FBF49A4F-82A3-43EB-8BA2-44D21233B114
        private static string GetTableName(string input)
        {
            var tableName = new char[input.Length];

            for (var i = 0; i < input.Length; i++)
            {
                tableName[i] = char.IsLetterOrDigit(input[i])
                    ? input[i]
                    : '_';
            }

            return new string(tableName);
        }

        private static string GetWorksheetName(object firstRow, Type collectionType)
        {
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

            if (worksheetNameProperty is object)
            {
                var worksheetNameAttribute = worksheetNameProperty.GetCustomAttribute<SpreadsheetTabNameAttribute>();
                var worksheetPropertyValue = worksheetNameProperty.GetValue(firstRow);

                if (!string.IsNullOrWhiteSpace(worksheetNameAttribute.FormatString))
                {
                    worksheetName = string.Format(worksheetNameAttribute.FormatString, worksheetPropertyValue);
                }
                else
                {
                    worksheetName = worksheetPropertyValue.ToString();
                }
            }
            else
            {
                var displayNameAttribute = collectionType.GetCustomAttribute<DisplayNameAttribute>();
                if (displayNameAttribute is object)
                {
                    worksheetName = displayNameAttribute.DisplayName;
                }
                else
                {
                    var displayAttribute = collectionType.GetCustomAttribute<DisplayAttribute>();
                    if (displayAttribute is object)
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

            var spreadsheetTableStyleAttribute = collectionType.GetCustomAttribute<SpreadsheetTableStyleAttribute>();
            if (spreadsheetTableStyleAttribute is object)
            {
                tableStyle = spreadsheetTableStyleAttribute.TableStyle;
            }

            return tableStyle;
        }

        private static string GetPropertyName(PropertyInfo property, Type collectionType)
        {
            var propertyName = property.Name;

            var displayNameAttribute = property.GetCustomAttribute<DisplayNameAttribute>();
            if (displayNameAttribute is object)
            {
                propertyName = displayNameAttribute.DisplayName;
            }
            else
            {
                var displayAttribute = property.GetCustomAttribute<DisplayAttribute>();
                if (displayAttribute is object)
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

            var displayFormatAttribute = property.GetCustomAttribute<DisplayFormatAttribute>();
            if (displayFormatAttribute is object)
            {
                if (!(inputValue is object) && !string.IsNullOrWhiteSpace(displayFormatAttribute.NullDisplayText))
                {
                    outputValue = displayFormatAttribute.NullDisplayText;
                }
                else if (inputValue is object)
                {
                    outputValue = string.Format(displayFormatAttribute.DataFormatString, inputValue);
                }
            }
            else if (inputValue is object)
            {
                if (property.PropertyType.IsValueType)
                {
                    // for value types, just output the raw value
                    outputValue = inputValue;
                }
                else
                {
                    var enumerable = (inputValue as IEnumerable<object>);
                    if (enumerable is object)
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
                if (!(collection is object) || !collection.Any())
                {
                    continue;
                }

                var firstRow = collection.First();
                var collectionType = firstRow.GetType();
                var properties = collectionType.GetProperties();
                var worksheetName = GetWorksheetName(firstRow, collectionType);
                var worksheet = package.Workbook.Worksheets[worksheetName];

                if (!(worksheet is object))
                {
                    continue;
                }

                // loop through the properties in the collection type
                // and see if any have a SpreadsheetLinkAttribute specified
                for (var prop = 1; prop <= properties.Count(); prop++)
                {
                    var property = properties.ElementAt(prop - 1);

                    if (property.GetCustomAttribute<SpreadsheetExcludeAttribute>() is object)
                    {
                        continue;
                    }

                    var attribute = property.GetCustomAttribute<SpreadsheetLinkAttribute>();

                    if (!(attribute is object))
                    {
                        // no SpreadsheetLinkAttribute for this property,
                        // so skip to the next property
                        continue;
                    }

                    // get the worksheet specified by the attribute
                    var linkSheet = package.Workbook.Worksheets[attribute.WorksheetName];

                    if (!(linkSheet is object))
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
                var worksheetCell = worksheet.Cells[$"{GetColumnLetter(worksheetColumnIndex)}{worksheetRow}"];
                var worksheetValue = worksheetCell.Value.ToString();

                // loop through the cells of the target worksheet column and see if any of the values
                // match the value of the current worksheet cell
                for (var linksheetRow = 1; linksheetRow <= linkSheet.Dimension.Rows; linksheetRow++)
                {
                    var linksheetValue = linkSheet.Cells[$"{linkColumn}{linksheetRow}"].Value.ToString();

                    if (worksheetValue.Equals(linksheetValue))
                    {
                        // we found a match! this is the link target,
                        // so add the hyperlink to the worksheet cell
                        // and stop searching for targets for this row
                        worksheetCell.Hyperlink = new ExcelHyperLink($"'{linkSheet.Name}'!{linkColumn}{linksheetRow}", worksheetValue.ToString());
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
