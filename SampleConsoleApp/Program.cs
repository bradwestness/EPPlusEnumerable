using EPPlusEnumerable;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.IO;
using System.Linq;

namespace SampleConsoleApp
{
    class Program
    {
        static void Main(string[] args)
        {
            var data = new List<IEnumerable<object>>();

            using (var db = new SampleDataContext())
            {
                data.Add(db.Users.OrderBy(x => x.Name).ToList());

                foreach (var grouping in db.Orders.OrderBy(x => x.Date).GroupBy(x => x.Date.Month))
                {
                    data.Add(grouping.ToList());
                }
            }

            var bytes = Spreadsheet.Create(data);
            File.WriteAllBytes("MySpreadsheet.xlsx", bytes);
        }
    }

    [DisplayName("Customers")]
    public class User
    {
        public string Name { get; set; }

        [SpreadsheetExclude]
        public string Password { get; set; }

        [SpreadsheetCellStyle(verticalAlignment: ExcelVerticalAlignment.Top, horizontalAlignment: ExcelHorizontalAlignment.Center, wrapText: true)]
        public string Address { get; set; }

        public string City { get; set; }

        [SpreadsheetCellStyle(bold: true, backgroundHtmlColor: "#00cc00", fontHtmlColor: "red")]
        public string Country { get; set; }

        public string Zip { get; set; }
    }

    [DisplayName("Orders"), SpreadsheetTableStyle(TableStyles.Medium16)]
    public class Order
    {
        public int Number { get; set; }

        public string Item { get; set; }

        [SpreadsheetLink("Customers", "Name")]
        public string Customer { get; set; }

        [DisplayFormat(DataFormatString = "{0:c}")]
        public decimal Price { get; set; }

        [SpreadsheetTabName(FormatString = "{0:MMMM yyyy}")]
        public DateTime Date { get; set; }
    }
}
