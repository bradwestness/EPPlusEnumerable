namespace EPPlusEnumerable.Tests
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Reflection;

    using Faker;
    using Faker.Generators;

    using OfficeOpenXml;

    using Xunit;

    public class UnitTests
    {
        [Fact]
        public void WorksheetNameAttribute()
        {
            // Arrange
            List<IGrouping<DateTime, SampleTestModel>> groupedData = UnitTests.GivenGroupedData();

            // Act
            ExcelPackage package = UnitTests.WhenExporting(groupedData);

            // Assert
            Assert.True(package.Workbook.Worksheets.Count == groupedData.Count);
            for (var i = 1; i < package.Workbook.Worksheets.Count; i++)
            {
                Assert.True(package.Workbook.Worksheets[i].Dimension.End.Row == groupedData[i - 1].Count() + 1);
                Assert.True(package.Workbook.Worksheets[i].Name == groupedData[i - 1].Key.ToString("MMMM yyyy"));
                UnitTests.ThenGeneratedWorksheetIsValid(package.Workbook.Worksheets[i], groupedData[i - 1].ToList());
            }
        }

        private static List<IGrouping<DateTime, SampleTestModel>> GivenGroupedData()
        {
            var rows = 1000;
            var dataGenerator = new Fake<SampleTestModel>();
            dataGenerator.SetProperty(x => x.DateTime, () => new DateTime(DateTime.UtcNow.Year, Numbers.Int(1, 13), 1));
            IList<SampleTestModel> data = dataGenerator.Generate(rows);
            List<IGrouping<DateTime, SampleTestModel>> groupedData = data.GroupBy(x => x.DateTime).ToList();
            return groupedData;
        }

        [Fact]
        public void
            GivenDataContainsPropertyMarkedWithExcludeAttributeWhenExportingThenGeneratedWorksheetDoesNotContainColumnForThisProperty
            ()
        {
            IList<SampleTestModelWithExcludedProperty> data =
                UnitTests.GivenDataContainingPropertyMarkedWithExcludeAttribute();
            ExcelPackage package = UnitTests.WhenExporting(data);

            UnitTests.ThenGeneratedWorksheetIsValid(package.Workbook.Worksheets.First(), data);
            UnitTests.ThenGeneratedWorksheetDoesNotContainColumnForThisProperty(
                package.Workbook.Worksheets.First(),
                data);
        }

        private static void ThenGeneratedWorksheetDoesNotContainColumnForThisProperty<T>(
            ExcelWorksheet worksheet,
            IList<T> data)
        {
            T first = data.First();
            Type parameterType = first.GetType();
            PropertyInfo[] properties = parameterType.GetProperties();

            foreach (PropertyInfo propertyInfo in properties)
            {
                var shouldExport = propertyInfo.GetCustomAttribute<SpreadsheetExcludeFromOutputAttribute>() != null;
                if (shouldExport)
                {
                    return;
                }

                var isEmpty =
                    !(from c in worksheet.Cells[1, 1, 1, 100] where c.Text != propertyInfo.Name select c).Any();
                Assert.False(
                    isEmpty,
                    string.Format(
                        "The generated worksheet contains column for excluded property:  \"{0}\"",
                        propertyInfo.Name));
            }
        }

        private static void ThenGeneratedWorksheetIsValid<T>(ExcelWorksheet worksheet, IList<T> data)
        {
            Assert.True(UnitTests.ValidWorksheet(worksheet, data));
        }

        private static ExcelPackage WhenExporting<T>(IList<T> data)
        {
            if (data is IEnumerable<IEnumerable<object>>)
            {
                var sets = (IEnumerable<IEnumerable<object>>)data;
                return Spreadsheet.CreatePackage(sets);
            }

            IEnumerable<object> list = data.Cast<object>();
            return Spreadsheet.CreatePackage(list);
        }

        private static IList<SampleTestModelWithExcludedProperty> GivenDataContainingPropertyMarkedWithExcludeAttribute(
            )
        {
            var rows = 100;
            var dataGenerator = new Fake<SampleTestModelWithExcludedProperty>();
            dataGenerator.SetProperty(x => x.Created, () => DateTime.UtcNow);
            dataGenerator.SetProperty(x => x.Name, () => Names.FullName());
            dataGenerator.SetProperty(x => x.Email, () => EmailAddresses.Generate(false, 10, 100));
            IList<SampleTestModelWithExcludedProperty> data = dataGenerator.Generate(rows);
            return data;
        }

        [Fact]
        public void NoRows()
        {
            // Arrange
            var data = new List<SampleTestModel>();

            // Act
            ExcelPackage package = Spreadsheet.CreatePackage(data);

            // Assert
            Assert.True(package.Workbook.Worksheets.Count == 0);
        }

        [Fact]
        public void ValidateData()
        {
            // Arrange
            var rows = 1000;
            var dataGenerator = new Fake<SampleTestModel>();
            IList<SampleTestModel> data = dataGenerator.Generate(rows);

            // Act
            ExcelPackage package = Spreadsheet.CreatePackage(data);

            // Assert
            Assert.True(package.Workbook.Worksheets.Count == 1);
            Assert.True(package.Workbook.Worksheets[1].Dimension.End.Row == rows + 1);
            Assert.True(UnitTests.ValidWorksheet(package.Workbook.Worksheets[1], data));
        }

        private static bool ValidWorksheet<T>(ExcelWorksheet excelWorksheet, IList<T> data)
        {
            for (var i = 2; i <= data.Count; i++)
            {
                var validRow = UnitTests.ValidRow(excelWorksheet, i, data[i - 2]);
                if (!validRow)
                {
                    return false;
                }
            }

            return true;
        }

        private static bool ValidRow<T>(ExcelWorksheet excelWorksheet, int row, T testModel)
        {
            List<PropertyInfo> properties = testModel.GetType().GetProperties().ToList();

            var count = 0;
            foreach (PropertyInfo propertyInfo in properties)
            {
                var shouldSkip = propertyInfo.GetCustomAttribute<SpreadsheetExcludeFromOutputAttribute>() != null;
                if (shouldSkip)
                {
                    continue;
                }

                count++;
                object value = excelWorksheet.Cells[row, count].Value;
                if (propertyInfo.GetValue(testModel).ToString() != value.ToString())
                {
                    return false;
                }
            }

            return true;
        }
    }

    internal class SampleTestModel
    {
        public int Id { get; set; }

        public string Name { get; set; }

        public string Email { get; set; }

        [SpreadsheetTabName(FormatString = "{0:MMMM yyyy}")]
        public DateTime DateTime { get; set; }
    }

    internal class SampleTestModelWithExcludedProperty
    {
        public int Id { get; set; }

        public string Name { get; set; }

        public string Email { get; set; }

        [SpreadsheetExcludeFromOutput]
        public DateTime Created { get; set; }
    }
}