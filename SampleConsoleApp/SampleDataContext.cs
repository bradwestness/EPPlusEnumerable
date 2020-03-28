using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;

namespace SampleConsoleApp
{
    public class SampleDataContext : IDisposable
    {
        #region Properties

        private static readonly Lazy<IList<Northwind.Customer>> _data = new Lazy<IList<Northwind.Customer>>(() =>
        {
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

            return new Northwind.NorthwindEntities(new Uri("https://services.odata.org/northwind/northwind.svc/"))
                .Customers
                .Expand("Orders")
                .Expand("Orders/Order_Details")
                .Expand("Orders/Order_Details/Product")
                .ToList();
        });

        public IEnumerable<User> Users { get; set; }

        public IEnumerable<Order> Orders { get; set; }

        #endregion

        #region Constructors

        public SampleDataContext()
        {
            IList<User> users = new List<User>();
            IList<Order> orders = new List<Order>();

            foreach (var customer in _data.Value)
            {
                if (customer.Orders.Any(x => x.OrderDate.HasValue && x.Order_Details.Any(y => y.Product is object)))
                {
                    users.Add(new User
                    {
                        Name = customer.ContactName,
                        Address = customer.Address,
                        City = customer.City,
                        Country = customer.Country,
                        Zip = customer.PostalCode
                    });

                    foreach (var order in customer.Orders.Where(x => x.OrderDate.HasValue && x.Order_Details.Any(y => y.Product is object)))
                    {
                        orders.Add(new Order
                        {
                            Customer = customer.ContactName,
                            Number = order.OrderID,
                            Date = order.OrderDate.Value,
                            Item = order.Order_Details.First().Product.ProductName,
                            Price = order.Order_Details.First().UnitPrice
                        });
                    }
                }
            }

            Users = users.OrderBy(x => x.Name).ToList();
            Orders = orders.OrderBy(x => x.Date).ToList();
        }

        #endregion

        #region IDisposable Implementation

        public void Dispose()
        {
            Users = null;
            Orders = null;
        }

        #endregion
    }
}