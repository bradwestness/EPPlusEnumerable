using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Drawing;

namespace EPPlusEnumerable
{
    /// <summary>
    /// Use this attribute to denote the cell style EPPlus should use for a given column on the worksheet.
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
    public class SpreadsheetCellStyleAttribute : Attribute
    {
        #region Properties

        public ExcelHorizontalAlignment HorizontalAlignment { get; set; }

        public ExcelVerticalAlignment VerticalAlignment { get; set; }

        public bool WrapText { get; set; }

        public bool Bold { get; set; }

        public string BackgroundHtmlColor { get; set; }

        public string FontHtmlColor { get; set; }

        #endregion

        #region Constructors

        /// <summary>
        /// Apply the given cell style options to the worksheet for this collection property.
        /// </summary>
        /// <param name="horizontalAlignment">The horizontal alignment to be applied to the cells for this property.</param>
        /// <param name="verticalAlignment">The vertical alignment to be applied to the cells for this property.</param>
        /// <param name="wrapText">Controls whether text will be wrapped in the cells for this property.</param>
        /// <param name="bold">Controls whether text will be bold in the cells for this property.</param>
        /// <param name="backgroundHtmlColor">Controls the background color of the cells for this property. Can be either a hex code ("#00cc00") or an HTML color name ("lime").</param>
        /// <param name="fontHtmlColor">Controls the font color of the cells for this property. Can be either a hex code ("#00cc00") or an HTML color name ("lime").</param>
        public SpreadsheetCellStyleAttribute(ExcelHorizontalAlignment horizontalAlignment = ExcelHorizontalAlignment.General, ExcelVerticalAlignment verticalAlignment = ExcelVerticalAlignment.Bottom, bool wrapText = false, bool bold = false, string backgroundHtmlColor = null, string fontHtmlColor = null)
        {
            HorizontalAlignment = horizontalAlignment;
            VerticalAlignment = verticalAlignment;
            WrapText = wrapText;
            Bold = bold;
            BackgroundHtmlColor = backgroundHtmlColor;
            FontHtmlColor = fontHtmlColor;
        }

        #endregion

        public void ApplyCellStyles(ExcelRange range)
        {
            if (!(range is object))
            {
                return;
            }

            range.Style.HorizontalAlignment = HorizontalAlignment;
            range.Style.VerticalAlignment = VerticalAlignment;
            range.Style.WrapText = WrapText;
            range.Style.Font.Bold = Bold;

            if (!string.IsNullOrWhiteSpace(BackgroundHtmlColor))
            {
                range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor.SetColor(HexToColor(BackgroundHtmlColor));
            }

            if (!string.IsNullOrWhiteSpace(FontHtmlColor))
            {
                range.Style.Font.Color.SetColor(HexToColor(FontHtmlColor));
            }
        }

        private Color HexToColor(string hexColor) =>
            ColorTranslator.FromHtml(hexColor);
    }
}
