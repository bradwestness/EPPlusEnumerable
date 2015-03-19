using EPPlusEnumerable;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SampleConsoleApp
{
    class Program
    {
        static void Main(string[] args)
        {
            //Example 1 - Tab name from class annotation
            var data = new List<IEnumerable<object>>();

            using (var db = new SampleDataContext())
            {
                data.Add(db.Users.OrderBy(x => x.Name).ToList());
                data.Add(db.Orders.OrderBy(x => x.Customer).ThenByDescending(x => x.Date).ToList());
            }


            var bytes = Spreadsheet.Create(data);
            File.WriteAllBytes("MySpreadsheet.xlsx", bytes);


            //Example 2 - With custom tab names
            var data2 = new Dictionary<string, IEnumerable<object>>();

            using (var db = new SampleDataContext())
            {
                data2.Add("Tab 1", db.Users.OrderBy(x => x.Name).ToList());
                data2.Add("Tab 2", db.Orders2.OrderBy(x => x.Customer).ThenByDescending(x => x.Date).ToList());
            }

            var bytes2 = Spreadsheet.Create(data2);
            File.WriteAllBytes("MySpreadsheet2.xlsx", bytes2);
        }
    }

    [DisplayName("Customers")]
    public class User
    {
        public string Name { get; set; }

        public string Address { get; set; }

        public string City { get; set; }

        public string State { get; set; }

        public string Zip { get; set; }
    }

    [DisplayName("Orders")]
    public class Order
    {
        public int? Number { get; set; }

        public string Item { get; set; }

        [SpreadsheetLink("Customers", "Name")]
        public string Customer { get; set; }

        [DisplayFormat(DataFormatString="{0:0.000}")]
        public decimal? Price { get; set; }

        [DisplayFormat(DataFormatString="{0:dd/MM/yyyy}")]
        public DateTime? Date { get; set; }
    }

    [DisplayName("Orders")]
    public class Order2
    {
        public int? Number { get; set; }

        public string Item { get; set; }

        [SpreadsheetLink("Tab 1", "Name")]
        public string Customer { get; set; }

        [DisplayFormat(DataFormatString = "{0:0.000}")]
        public decimal? Price { get; set; }

        [DisplayFormat(DataFormatString = "{0:dd/MM/yyyy}")]
        public DateTime? Date { get; set; }
    }
}
