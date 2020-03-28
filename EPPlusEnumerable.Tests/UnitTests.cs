namespace EPPlusEnumerable.Tests
{
    using Faker;
    using Faker.Generators;
    using OfficeOpenXml;
    using OfficeOpenXml.Style;
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using Xunit;

    public class UnitTests
    {
        [Fact]
        public void WorksheetNameAttribute()
        {
            // Arrange
            int rows = 1000;
            var dataGenerator = new Fake<Grid>();
            dataGenerator.SetProperty(x => x.DateTime, () => new DateTime(DateTime.UtcNow.Year, Numbers.Int(1, 13), 1));
            var data = dataGenerator.Generate(rows);
            var groupedData = data.GroupBy(x => x.DateTime).ToList();

            // Act
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var package = Spreadsheet.CreatePackage(groupedData);

            // Assert
            Assert.True(package.Workbook.Worksheets.Count == groupedData.Count);
            for (int i = 1; i < package.Workbook.Worksheets.Count; i++)
            {
                Assert.True(package.Workbook.Worksheets[i].Dimension.End.Row == groupedData[i - 1].Count() + 1);
                Assert.True(package.Workbook.Worksheets[i].Name == groupedData[i - 1].Key.ToString("MMMM yyyy"));
                Assert.True(ValidWorksheet(package.Workbook.Worksheets[i], groupedData[i - 1].ToList()));
            }

        }

        [Fact]
        public void NoRows()
        {
            // Arrange
            var data = new List<Grid>();

            // Act
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var package = Spreadsheet.CreatePackage(data);

            // Assert
            Assert.True(package.Workbook.Worksheets.Count == 0);

        }

        [Fact]
        public void ValidateData()
        {
            // Arrange
            int rows = 1000;
            var dataGenerator = new Fake<Grid>();
            var data = dataGenerator.Generate(rows);

            // Act
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var package = Spreadsheet.CreatePackage(data);

            // Assert
            Assert.True(package.Workbook.Worksheets.Count == 1);
            Assert.True(package.Workbook.Worksheets[1].Dimension.End.Row == rows + 1);
            Assert.True(ValidWorksheet(package.Workbook.Worksheets[1], data));
        }

        private static bool ValidWorksheet(ExcelWorksheet excelWorksheet, IList<Grid> data)
        {
            for (int i = 2; i <= data.Count; i++)
            {
                var validRow = ValidRow(excelWorksheet, i, data[i - 2]);
                if (!validRow)
                {
                    return false;
                }
            }

            return true;
        }

        private static bool ValidRow(ExcelWorksheet excelWorksheet, int row, Grid grid)
        {
            var properties = grid.GetType().GetProperties().ToList();

            foreach (var propertyInfo in properties)
            {
                var value = excelWorksheet.Cells[row, properties.IndexOf(propertyInfo) + 1].Value;
                if (propertyInfo.GetValue(grid).ToString() != value.ToString())
                {
                    return false;
                }
            }

            return true;
        }
    }

    internal class Grid
    {
        public int Id { get; set; }

        public string Name { get; set; }

        [SpreadsheetCellStyle(horizontalAlignment: ExcelHorizontalAlignment.Center, wrapText: true)]
        public string Email { get; set; }

        [SpreadsheetTabName(FormatString = "{0:MMMM yyyy}")]
        public DateTime DateTime { get; set; }
    }
}