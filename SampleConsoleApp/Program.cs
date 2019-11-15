using EPPlusEnumerable;
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

                foreach(var grouping in db.Orders.OrderBy(x => x.Date).GroupBy(x => x.Date.Month))
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

        public string Address { get; set; }

        public string City { get; set; }

        public string Country { get; set; }

        public string Zip { get; set; }
    }
    
    [MetadataType(typeof(OrdersMetadata))]
    [DisplayName("Orders"), SpreadsheetTableStyle(TableStyles.Medium16)]
    public partial class Order
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
    
    // Metadata class for Orders. The new GetMetaAttribute extension method will find this and crawl it for attributes.
    public class OrdersMetadata
    {
        [Display(Name = "Order #")]
        public int Number { get; set; }
        [DisplayName("Item Price")]
        public decimal Price { get; set; }
        [DisplayName("Customer Name")]
        public string Customer { get; set; }
    }
}
