using System;
using System.Collections.Generic;

namespace SampleConsoleApp
{
    public class SampleDataContext : IDisposable
    {
        #region Properties

        public IEnumerable<User> Users { get; set; }

        public IEnumerable<Order> Orders { get; set; }

        public IEnumerable<Order2> Orders2 { get; set; }

        #endregion

        #region Constructors

        public SampleDataContext()
        {
            Users = GetUsers();
            Orders = GetOrders();
            Orders2 = GetOrders2();
        }

        #endregion

        #region Private Methods

        private static IEnumerable<User> GetUsers()
        {
            return new[] {
                new User{ Name = "Jackson Turner", Address = "123 Main St", City = "New York", State = "NY", Zip = "11226"},
                new User{ Name = "Megan Perry", Address = "123 Oak St", City = "Los Angeles", State = "CA", Zip = "90011"},
                new User{ Name = "Ryan Harris", Address = "123 Pine St", City = "Chicago", State = "IL", Zip = "60629"},
                new User{ Name = "Mason Edwards", Address = "123 Maple St", City = "Houston", State = "TX", Zip = "77084"},
                new User{ Name = "Noah Jenkins", Address = "123 Cedar St", City = "Philadelphia", State = "PA", Zip = "19103"},
                new User{ Name = "Stephanie Hayes", Address = "123 Elm St", City = "Phoenix", State = "AZ", Zip = "85001"},
                new User{ Name = "Caleb Scott", Address = "123 View St", City = "San Antonio", State = "TX", Zip = "78201"},
                new User{ Name = "Morgan Wood", Address = "123 Washington St", City = "Dallas", State = "TX", Zip = "75217"},
                new User{ Name = "James Parker", Address = "123 Lake St", City = "San Jose", State = "CA", Zip = "95113"},
                new User{ Name = "Austin Jackson", Address = "123 Hill St", City = "Indianapolis", State = "IN", Zip = "46268"},
            };
        }

        private static IEnumerable<Order> GetOrders()
        {
            return new[] {
                new Order{ Number = 1, Item = "Water Filter", Customer = "Jackson Turner", Price = (decimal)13.69, Date= DateTime.Now},
                new Order{ Number = 2, Item = "Digital Scale", Customer = "Jackson Turner", Price = (decimal)13.00, Date= DateTime.Now.AddDays(-1) },
                new Order{ Number = 3, Item = "Ceramic Heater", Customer = "Jackson Turner", Price = (decimal)24.97, Date= DateTime.Now.AddDays(-2)},
                new Order{ Number = 4, Item = "Humidifier", Customer = "Jackson Turner", Price = (decimal)88.79, Date= DateTime.Now.AddDays(-3)},
                new Order{ Number = 5, Item = "Coffee Filters", Customer = "Jackson Turner", Price = (decimal)7.79, Date= DateTime.Now.AddDays(-4)},
                new Order{ Number = 6, Item = "Crock Pot", Customer = "Jackson Turner", Price = (decimal)23.99, Date= DateTime.Now.AddDays(-5)},
                new Order{ Number = 7, Item = "Vacuum Cleaner", Customer = "Jackson Turner", Price = (decimal)79.00, Date= DateTime.Now.AddDays(-6)},
                new Order{ Number = 8, Item = "Coffee Pot", Customer = "Jackson Turner", Price = (decimal)24.92, Date= DateTime.Now.AddDays(-7)},
                new Order{ Number = 9, Item = "TV Stand", Customer = "Jackson Turner", Price = (decimal)25.77, Date= DateTime.Now.AddDays(-8)},
                new Order{ Number = 10, Item = "Electric Steam Mop", Customer = "Jackson Turner", Price = (decimal)79.99, Date= DateTime.Now.AddDays(-9)},
                
                new Order{ Number = 1, Item = "Water Filter", Customer = "Megan Perry", Price = (decimal)13.69, Date= DateTime.Now},
                new Order{ Number = 2, Item = "Digital Scale", Customer = "Megan Perry", Price = (decimal)13.00, Date= DateTime.Now.AddDays(-1) },
                new Order{ Number = 3, Item = "Ceramic Heater", Customer = "Megan Perry", Price = (decimal)24.97, Date= DateTime.Now.AddDays(-2)},
                new Order{ Number = 4, Item = "Humidifier", Customer = "Megan Perry", Price = (decimal)88.79, Date= DateTime.Now.AddDays(-3)},
                new Order{ Number = 5, Item = "Coffee Filters", Customer = "Megan Perry", Price = (decimal)7.79, Date= DateTime.Now.AddDays(-4)},
                new Order{ Number = 6, Item = "Crock Pot", Customer = "Megan Perry", Price = (decimal)23.99, Date= DateTime.Now.AddDays(-5)},
                new Order{ Number = 7, Item = "Vacuum Cleaner", Customer = "Megan Perry", Price = (decimal)79.00, Date= DateTime.Now.AddDays(-6)},
                new Order{ Number = 8, Item = "Coffee Pot", Customer = "Megan Perry", Price = (decimal)24.92, Date= DateTime.Now.AddDays(-7)},
                new Order{ Number = 9, Item = "TV Stand", Customer = "Megan Perry", Price = (decimal)25.77, Date= DateTime.Now.AddDays(-8)},
                new Order{ Number = 10, Item = "Electric Steam Mop", Customer = "Megan Perry", Price = (decimal)79.99, Date= DateTime.Now.AddDays(-9)},
                
                new Order{ Number = 1, Item = "Water Filter", Customer = "Ryan Harris", Price = (decimal)13.69, Date= DateTime.Now},
                new Order{ Number = 2, Item = "Digital Scale", Customer = "Ryan Harris", Price = (decimal)13.00, Date= DateTime.Now.AddDays(-1) },
                new Order{ Number = 3, Item = "Ceramic Heater", Customer = "Ryan Harris", Price = (decimal)24.97, Date= DateTime.Now.AddDays(-2)},
                new Order{ Number = 4, Item = "Humidifier", Customer = "Ryan Harris", Price = (decimal)88.79, Date= DateTime.Now.AddDays(-3)},
                new Order{ Number = 5, Item = "Coffee Filters", Customer = "Ryan Harris", Price = (decimal)7.79, Date= DateTime.Now.AddDays(-4)},
                new Order{ Number = 6, Item = "Crock Pot", Customer = "Ryan Harris", Price = (decimal)23.99, Date= DateTime.Now.AddDays(-5)},
                new Order{ Number = 7, Item = "Vacuum Cleaner", Customer = "Ryan Harris", Price = (decimal)79.00, Date= DateTime.Now.AddDays(-6)},
                new Order{ Number = 8, Item = "Coffee Pot", Customer = "Ryan Harris", Price = (decimal)24.92, Date= DateTime.Now.AddDays(-7)},
                new Order{ Number = 9, Item = "TV Stand", Customer = "Ryan Harris", Price = (decimal)25.77, Date= DateTime.Now.AddDays(-8)},
                new Order{ Number = 10, Item = "Electric Steam Mop", Customer = "Ryan Harris", Price = (decimal)79.99, Date= DateTime.Now.AddDays(-9)},
                
                new Order{ Number = 1, Item = "Water Filter", Customer = "Mason Edwards", Price = (decimal)13.69, Date= DateTime.Now},
                new Order{ Number = 2, Item = "Digital Scale", Customer = "Mason Edwards", Price = (decimal)13.00, Date= DateTime.Now.AddDays(-1) },
                new Order{ Number = 3, Item = "Ceramic Heater", Customer = "Mason Edwards", Price = (decimal)24.97, Date= DateTime.Now.AddDays(-2)},
                new Order{ Number = 4, Item = "Humidifier", Customer = "Mason Edwards", Price = (decimal)88.79, Date= DateTime.Now.AddDays(-3)},
                new Order{ Number = 5, Item = "Coffee Filters", Customer = "Mason Edwards", Price = (decimal)7.79, Date= DateTime.Now.AddDays(-4)},
                new Order{ Number = 6, Item = "Crock Pot", Customer = "Mason Edwards", Price = (decimal)23.99, Date= DateTime.Now.AddDays(-5)},
                new Order{ Number = 7, Item = "Vacuum Cleaner", Customer = "Mason Edwards", Price = (decimal)79.00, Date= DateTime.Now.AddDays(-6)},
                new Order{ Number = 8, Item = "Coffee Pot", Customer = "Mason Edwards", Price = (decimal)24.92, Date= DateTime.Now.AddDays(-7)},
                new Order{ Number = 9, Item = "TV Stand", Customer = "Mason Edwards", Price = (decimal)25.77, Date= DateTime.Now.AddDays(-8)},
                new Order{ Number = 10, Item = "Electric Steam Mop", Customer = "Mason Edwards", Price = (decimal)79.99, Date= DateTime.Now.AddDays(-9)},
                
                new Order{ Number = 1, Item = "Water Filter", Customer = "Noah Jenkins", Price = (decimal)13.69, Date= DateTime.Now},
                new Order{ Number = 2, Item = "Digital Scale", Customer = "Noah Jenkins", Price = (decimal)13.00, Date= DateTime.Now.AddDays(-1) },
                new Order{ Number = 3, Item = "Ceramic Heater", Customer = "Noah Jenkins", Price = (decimal)24.97, Date= DateTime.Now.AddDays(-2)},
                new Order{ Number = 4, Item = "Humidifier", Customer = "Noah Jenkins", Price = (decimal)88.79, Date= DateTime.Now.AddDays(-3)},
                new Order{ Number = 5, Item = "Coffee Filters", Customer = "Noah Jenkins", Price = (decimal)7.79, Date= DateTime.Now.AddDays(-4)},
                new Order{ Number = 6, Item = "Crock Pot", Customer = "Noah Jenkins", Price = (decimal)23.99, Date= DateTime.Now.AddDays(-5)},
                new Order{ Number = 7, Item = "Vacuum Cleaner", Customer = "Noah Jenkins", Price = (decimal)79.00, Date= DateTime.Now.AddDays(-6)},
                new Order{ Number = 8, Item = "Coffee Pot", Customer = "Noah Jenkins", Price = (decimal)24.92, Date= DateTime.Now.AddDays(-7)},
                new Order{ Number = 9, Item = "TV Stand", Customer = "Noah Jenkins", Price = (decimal)25.77, Date= DateTime.Now.AddDays(-8)},
                new Order{ Number = 10, Item = "Electric Steam Mop", Customer = "Noah Jenkins", Price = (decimal)79.99, Date= DateTime.Now.AddDays(-9)},
                
                new Order{ Number = 1, Item = "Water Filter", Customer = "Stephanie Hayes", Price = (decimal)13.69, Date= DateTime.Now},
                new Order{ Number = 2, Item = "Digital Scale", Customer = "Stephanie Hayes", Price = (decimal)13.00, Date= DateTime.Now.AddDays(-1) },
                new Order{ Number = 3, Item = "Ceramic Heater", Customer = "Stephanie Hayes", Price = (decimal)24.97, Date= DateTime.Now.AddDays(-2)},
                new Order{ Number = 4, Item = "Humidifier", Customer = "Stephanie Hayes", Price = (decimal)88.79, Date= DateTime.Now.AddDays(-3)},
                new Order{ Number = 5, Item = "Coffee Filters", Customer = "Stephanie Hayes", Price = (decimal)7.79, Date= DateTime.Now.AddDays(-4)},
                new Order{ Number = 6, Item = "Crock Pot", Customer = "Stephanie Hayes", Price = (decimal)23.99, Date= DateTime.Now.AddDays(-5)},
                new Order{ Number = 7, Item = "Vacuum Cleaner", Customer = "Stephanie Hayes", Price = (decimal)79.00, Date= DateTime.Now.AddDays(-6)},
                new Order{ Number = 8, Item = "Coffee Pot", Customer = "Stephanie Hayes", Price = (decimal)24.92, Date= DateTime.Now.AddDays(-7)},
                new Order{ Number = 9, Item = "TV Stand", Customer = "Stephanie Hayes", Price = (decimal)25.77, Date= DateTime.Now.AddDays(-8)},
                new Order{ Number = 10, Item = "Electric Steam Mop", Customer = "Stephanie Hayes", Price = (decimal)79.99, Date= DateTime.Now.AddDays(-9)},
                
                new Order{ Number = 1, Item = "Water Filter", Customer = "Caleb Scott", Price = (decimal)13.69, Date= DateTime.Now},
                new Order{ Number = 2, Item = "Digital Scale", Customer = "Caleb Scott", Price = (decimal)13.00, Date= DateTime.Now.AddDays(-1) },
                new Order{ Number = 3, Item = "Ceramic Heater", Customer = "Caleb Scott", Price = (decimal)24.97, Date= DateTime.Now.AddDays(-2)},
                new Order{ Number = 4, Item = "Humidifier", Customer = "Caleb Scott", Price = (decimal)88.79, Date= DateTime.Now.AddDays(-3)},
                new Order{ Number = 5, Item = "Coffee Filters", Customer = "Caleb Scott", Price = (decimal)7.79, Date= DateTime.Now.AddDays(-4)},
                new Order{ Number = 6, Item = "Crock Pot", Customer = "Caleb Scott", Price = (decimal)23.99, Date= DateTime.Now.AddDays(-5)},
                new Order{ Number = 7, Item = "Vacuum Cleaner", Customer = "Caleb Scott", Price = (decimal)79.00, Date= DateTime.Now.AddDays(-6)},
                new Order{ Number = 8, Item = "Coffee Pot", Customer = "Caleb Scott", Price = (decimal)24.92, Date= DateTime.Now.AddDays(-7)},
                new Order{ Number = 9, Item = "TV Stand", Customer = "Caleb Scott", Price = (decimal)25.77, Date= DateTime.Now.AddDays(-8)},
                new Order{ Number = 10, Item = "Electric Steam Mop", Customer = "Caleb Scott", Price = (decimal)79.99, Date= DateTime.Now.AddDays(-9)},
                
                new Order{ Number = 1, Item = "Water Filter", Customer = "Morgan Wood", Price = (decimal)13.69, Date= DateTime.Now},
                new Order{ Number = 2, Item = "Digital Scale", Customer = "Morgan Wood", Price = (decimal)13.00, Date= DateTime.Now.AddDays(-1) },
                new Order{ Number = 3, Item = "Ceramic Heater", Customer = "Morgan Wood", Price = (decimal)24.97, Date= DateTime.Now.AddDays(-2)},
                new Order{ Number = 4, Item = "Humidifier", Customer = "Morgan Wood", Price = (decimal)88.79, Date= DateTime.Now.AddDays(-3)},
                new Order{ Number = 5, Item = "Coffee Filters", Customer = "Morgan Wood", Price = (decimal)7.79, Date= DateTime.Now.AddDays(-4)},
                new Order{ Number = 6, Item = "Crock Pot", Customer = "Morgan Wood", Price = (decimal)23.99, Date= DateTime.Now.AddDays(-5)},
                new Order{ Number = 7, Item = "Vacuum Cleaner", Customer = "Morgan Wood", Price = (decimal)79.00, Date= DateTime.Now.AddDays(-6)},
                new Order{ Number = 8, Item = "Coffee Pot", Customer = "Morgan Wood", Price = (decimal)24.92, Date= DateTime.Now.AddDays(-7)},
                new Order{ Number = 9, Item = "TV Stand", Customer = "Morgan Wood", Price = (decimal)25.77, Date= DateTime.Now.AddDays(-8)},
                new Order{ Number = 10, Item = "Electric Steam Mop", Customer = "Morgan Wood", Price = (decimal)79.99, Date= DateTime.Now.AddDays(-9)},
                
                new Order{ Number = 1, Item = "Water Filter", Customer = "James Parker", Price = (decimal)13.69, Date= DateTime.Now},
                new Order{ Number = 2, Item = "Digital Scale", Customer = "James Parker", Price = (decimal)13.00, Date= DateTime.Now.AddDays(-1) },
                new Order{ Number = 3, Item = "Ceramic Heater", Customer = "James Parker", Price = (decimal)24.97, Date= DateTime.Now.AddDays(-2)},
                new Order{ Number = 4, Item = "Humidifier", Customer = "James Parker", Price = (decimal)88.79, Date= DateTime.Now.AddDays(-3)},
                new Order{ Number = 5, Item = "Coffee Filters", Customer = "James Parker", Price = (decimal)7.79, Date= DateTime.Now.AddDays(-4)},
                new Order{ Number = 6, Item = "Crock Pot", Customer = "James Parker", Price = (decimal)23.99, Date= DateTime.Now.AddDays(-5)},
                new Order{ Number = 7, Item = "Vacuum Cleaner", Customer = "James Parker", Price = (decimal)79.00, Date= DateTime.Now.AddDays(-6)},
                new Order{ Number = 8, Item = "Coffee Pot", Customer = "James Parker", Price = (decimal)24.92, Date= DateTime.Now.AddDays(-7)},
                new Order{ Number = 9, Item = "TV Stand", Customer = "James Parker", Price = (decimal)25.77, Date= DateTime.Now.AddDays(-8)},
                new Order{ Number = 10, Item = "Electric Steam Mop", Customer = "James Parker", Price = (decimal)79.99, Date= DateTime.Now.AddDays(-9)},
                
                new Order{ Number = 1, Item = "Water Filter", Customer = "Austin Jackson", Price = (decimal)13.69, Date= DateTime.Now},
                new Order{ Number = 2, Item = "Digital Scale", Customer = "Austin Jackson", Price = (decimal)13.00, Date= DateTime.Now.AddDays(-1) },
                new Order{ Number = 3, Item = "Ceramic Heater", Customer = "Austin Jackson", Price = (decimal)24.97, Date= DateTime.Now.AddDays(-2)},
                new Order{ Number = 4, Item = "Humidifier", Customer = "Austin Jackson", Price = (decimal)88.79, Date= DateTime.Now.AddDays(-3)},
                new Order{ Number = 5, Item = "Coffee Filters", Customer = "Austin Jackson", Price = (decimal)7.79, Date= DateTime.Now.AddDays(-4)},
                new Order{ Number = 6, Item = "Crock Pot", Customer = "Austin Jackson", Price = (decimal)23.99, Date= DateTime.Now.AddDays(-5)},
                new Order{ Number = 7, Item = "Vacuum Cleaner", Customer = "Austin Jackson", Price = (decimal)79.00, Date= DateTime.Now.AddDays(-6)},
                new Order{ Number = 8, Item = "Coffee Pot", Customer = "Austin Jackson", Price = (decimal)24.92, Date= DateTime.Now.AddDays(-7)},
                new Order{ Number = 9, Item = "TV Stand", Customer = "Austin Jackson", Price = (decimal)25.77, Date= DateTime.Now.AddDays(-8)},
                new Order{ Number = 10, Item = "Electric Steam Mop", Customer = "Austin Jackson", Price = (decimal)79.99, Date= DateTime.Now.AddDays(-9)},
            };
        }

        private static IEnumerable<Order2> GetOrders2()
        {
            return new[] {
                new Order2{ Number = 1, Item = "Water Filter", Customer = "Jackson Turner", Price = (decimal)13.69, Date= DateTime.Now},
                new Order2{ Number = 2, Item = "Digital Scale", Customer = "Jackson Turner", Price = (decimal)13.00, Date= DateTime.Now.AddDays(-1) },
                new Order2{ Number = 3, Item = "Ceramic Heater", Customer = "Jackson Turner", Price = (decimal)24.97, Date= DateTime.Now.AddDays(-2)},
                new Order2{ Number = 4, Item = "Humidifier", Customer = "Jackson Turner", Price = (decimal)88.79, Date= DateTime.Now.AddDays(-3)},
                new Order2{ Number = 5, Item = "Coffee Filters", Customer = "Jackson Turner", Price = (decimal)7.79, Date= DateTime.Now.AddDays(-4)},
                new Order2{ Number = 6, Item = "Crock Pot", Customer = "Jackson Turner", Price = (decimal)23.99, Date= DateTime.Now.AddDays(-5)},
                new Order2{ Number = 7, Item = "Vacuum Cleaner", Customer = "Jackson Turner", Price = (decimal)79.00, Date= DateTime.Now.AddDays(-6)},
                new Order2{ Number = 8, Item = "Coffee Pot", Customer = "Jackson Turner", Price = (decimal)24.92, Date= DateTime.Now.AddDays(-7)},
                new Order2{ Number = 9, Item = "TV Stand", Customer = "Jackson Turner", Price = (decimal)25.77, Date= DateTime.Now.AddDays(-8)},
                new Order2{ Number = 10, Item = "Electric Steam Mop", Customer = "Jackson Turner", Price = (decimal)79.99, Date= DateTime.Now.AddDays(-9)},
                         
                new Order2{ Number = 1, Item = "Water Filter", Customer = "Megan Perry", Price = (decimal)13.69, Date= DateTime.Now},
                new Order2{ Number = 2, Item = "Digital Scale", Customer = "Megan Perry", Price = (decimal)13.00, Date= DateTime.Now.AddDays(-1) },
                new Order2{ Number = 3, Item = "Ceramic Heater", Customer = "Megan Perry", Price = (decimal)24.97, Date= DateTime.Now.AddDays(-2)},
                new Order2{ Number = 4, Item = "Humidifier", Customer = "Megan Perry", Price = (decimal)88.79, Date= DateTime.Now.AddDays(-3)},
                new Order2{ Number = 5, Item = "Coffee Filters", Customer = "Megan Perry", Price = (decimal)7.79, Date= DateTime.Now.AddDays(-4)},
                new Order2{ Number = 6, Item = "Crock Pot", Customer = "Megan Perry", Price = (decimal)23.99, Date= DateTime.Now.AddDays(-5)},
                new Order2{ Number = 7, Item = "Vacuum Cleaner", Customer = "Megan Perry", Price = (decimal)79.00, Date= DateTime.Now.AddDays(-6)},
                new Order2{ Number = 8, Item = "Coffee Pot", Customer = "Megan Perry", Price = (decimal)24.92, Date= DateTime.Now.AddDays(-7)},
                new Order2{ Number = 9, Item = "TV Stand", Customer = "Megan Perry", Price = (decimal)25.77, Date= DateTime.Now.AddDays(-8)},
                new Order2{ Number = 10, Item = "Electric Steam Mop", Customer = "Megan Perry", Price = (decimal)79.99, Date= DateTime.Now.AddDays(-9)},
                         
                new Order2{ Number = 1, Item = "Water Filter", Customer = "Ryan Harris", Price = (decimal)13.69, Date= DateTime.Now},
                new Order2{ Number = 2, Item = "Digital Scale", Customer = "Ryan Harris", Price = (decimal)13.00, Date= DateTime.Now.AddDays(-1) },
                new Order2{ Number = 3, Item = "Ceramic Heater", Customer = "Ryan Harris", Price = (decimal)24.97, Date= DateTime.Now.AddDays(-2)},
                new Order2{ Number = 4, Item = "Humidifier", Customer = "Ryan Harris", Price = (decimal)88.79, Date= DateTime.Now.AddDays(-3)},
                new Order2{ Number = 5, Item = "Coffee Filters", Customer = "Ryan Harris", Price = (decimal)7.79, Date= DateTime.Now.AddDays(-4)},
                new Order2{ Number = 6, Item = "Crock Pot", Customer = "Ryan Harris", Price = (decimal)23.99, Date= DateTime.Now.AddDays(-5)},
                new Order2{ Number = 7, Item = "Vacuum Cleaner", Customer = "Ryan Harris", Price = (decimal)79.00, Date= DateTime.Now.AddDays(-6)},
                new Order2{ Number = 8, Item = "Coffee Pot", Customer = "Ryan Harris", Price = (decimal)24.92, Date= DateTime.Now.AddDays(-7)},
                new Order2{ Number = 9, Item = "TV Stand", Customer = "Ryan Harris", Price = (decimal)25.77, Date= DateTime.Now.AddDays(-8)},
                new Order2{ Number = 10, Item = "Electric Steam Mop", Customer = "Ryan Harris", Price = (decimal)79.99, Date= DateTime.Now.AddDays(-9)},
                         
                new Order2{ Number = 1, Item = "Water Filter", Customer = "Mason Edwards", Price = (decimal)13.69, Date= DateTime.Now},
                new Order2{ Number = 2, Item = "Digital Scale", Customer = "Mason Edwards", Price = (decimal)13.00, Date= DateTime.Now.AddDays(-1) },
                new Order2{ Number = 3, Item = "Ceramic Heater", Customer = "Mason Edwards", Price = (decimal)24.97, Date= DateTime.Now.AddDays(-2)},
                new Order2{ Number = 4, Item = "Humidifier", Customer = "Mason Edwards", Price = (decimal)88.79, Date= DateTime.Now.AddDays(-3)},
                new Order2{ Number = 5, Item = "Coffee Filters", Customer = "Mason Edwards", Price = (decimal)7.79, Date= DateTime.Now.AddDays(-4)},
                new Order2{ Number = 6, Item = "Crock Pot", Customer = "Mason Edwards", Price = (decimal)23.99, Date= DateTime.Now.AddDays(-5)},
                new Order2{ Number = 7, Item = "Vacuum Cleaner", Customer = "Mason Edwards", Price = (decimal)79.00, Date= DateTime.Now.AddDays(-6)},
                new Order2{ Number = 8, Item = "Coffee Pot", Customer = "Mason Edwards", Price = (decimal)24.92, Date= DateTime.Now.AddDays(-7)},
                new Order2{ Number = 9, Item = "TV Stand", Customer = "Mason Edwards", Price = (decimal)25.77, Date= DateTime.Now.AddDays(-8)},
                new Order2{ Number = 10, Item = "Electric Steam Mop", Customer = "Mason Edwards", Price = (decimal)79.99, Date= DateTime.Now.AddDays(-9)},
                         
                new Order2{ Number = 1, Item = "Water Filter", Customer = "Noah Jenkins", Price = (decimal)13.69, Date= DateTime.Now},
                new Order2{ Number = 2, Item = "Digital Scale", Customer = "Noah Jenkins", Price = (decimal)13.00, Date= DateTime.Now.AddDays(-1) },
                new Order2{ Number = 3, Item = "Ceramic Heater", Customer = "Noah Jenkins", Price = (decimal)24.97, Date= DateTime.Now.AddDays(-2)},
                new Order2{ Number = 4, Item = "Humidifier", Customer = "Noah Jenkins", Price = (decimal)88.79, Date= DateTime.Now.AddDays(-3)},
                new Order2{ Number = 5, Item = "Coffee Filters", Customer = "Noah Jenkins", Price = (decimal)7.79, Date= DateTime.Now.AddDays(-4)},
                new Order2{ Number = 6, Item = "Crock Pot", Customer = "Noah Jenkins", Price = (decimal)23.99, Date= DateTime.Now.AddDays(-5)},
                new Order2{ Number = 7, Item = "Vacuum Cleaner", Customer = "Noah Jenkins", Price = (decimal)79.00, Date= DateTime.Now.AddDays(-6)},
                new Order2{ Number = 8, Item = "Coffee Pot", Customer = "Noah Jenkins", Price = (decimal)24.92, Date= DateTime.Now.AddDays(-7)},
                new Order2{ Number = 9, Item = "TV Stand", Customer = "Noah Jenkins", Price = (decimal)25.77, Date= DateTime.Now.AddDays(-8)},
                new Order2{ Number = 10, Item = "Electric Steam Mop", Customer = "Noah Jenkins", Price = (decimal)79.99, Date= DateTime.Now.AddDays(-9)},
                         
                new Order2{ Number = 1, Item = "Water Filter", Customer = "Stephanie Hayes", Price = (decimal)13.69, Date= DateTime.Now},
                new Order2{ Number = 2, Item = "Digital Scale", Customer = "Stephanie Hayes", Price = (decimal)13.00, Date= DateTime.Now.AddDays(-1) },
                new Order2{ Number = 3, Item = "Ceramic Heater", Customer = "Stephanie Hayes", Price = (decimal)24.97, Date= DateTime.Now.AddDays(-2)},
                new Order2{ Number = 4, Item = "Humidifier", Customer = "Stephanie Hayes", Price = (decimal)88.79, Date= DateTime.Now.AddDays(-3)},
                new Order2{ Number = 5, Item = "Coffee Filters", Customer = "Stephanie Hayes", Price = (decimal)7.79, Date= DateTime.Now.AddDays(-4)},
                new Order2{ Number = 6, Item = "Crock Pot", Customer = "Stephanie Hayes", Price = (decimal)23.99, Date= DateTime.Now.AddDays(-5)},
                new Order2{ Number = 7, Item = "Vacuum Cleaner", Customer = "Stephanie Hayes", Price = (decimal)79.00, Date= DateTime.Now.AddDays(-6)},
                new Order2{ Number = 8, Item = "Coffee Pot", Customer = "Stephanie Hayes", Price = (decimal)24.92, Date= DateTime.Now.AddDays(-7)},
                new Order2{ Number = 9, Item = "TV Stand", Customer = "Stephanie Hayes", Price = (decimal)25.77, Date= DateTime.Now.AddDays(-8)},
                new Order2{ Number = 10, Item = "Electric Steam Mop", Customer = "Stephanie Hayes", Price = (decimal)79.99, Date= DateTime.Now.AddDays(-9)},
                         
                new Order2{ Number = 1, Item = "Water Filter", Customer = "Caleb Scott", Price = (decimal)13.69, Date= DateTime.Now},
                new Order2{ Number = 2, Item = "Digital Scale", Customer = "Caleb Scott", Price = (decimal)13.00, Date= DateTime.Now.AddDays(-1) },
                new Order2{ Number = 3, Item = "Ceramic Heater", Customer = "Caleb Scott", Price = (decimal)24.97, Date= DateTime.Now.AddDays(-2)},
                new Order2{ Number = 4, Item = "Humidifier", Customer = "Caleb Scott", Price = (decimal)88.79, Date= DateTime.Now.AddDays(-3)},
                new Order2{ Number = 5, Item = "Coffee Filters", Customer = "Caleb Scott", Price = (decimal)7.79, Date= DateTime.Now.AddDays(-4)},
                new Order2{ Number = 6, Item = "Crock Pot", Customer = "Caleb Scott", Price = (decimal)23.99, Date= DateTime.Now.AddDays(-5)},
                new Order2{ Number = 7, Item = "Vacuum Cleaner", Customer = "Caleb Scott", Price = (decimal)79.00, Date= DateTime.Now.AddDays(-6)},
                new Order2{ Number = 8, Item = "Coffee Pot", Customer = "Caleb Scott", Price = (decimal)24.92, Date= DateTime.Now.AddDays(-7)},
                new Order2{ Number = 9, Item = "TV Stand", Customer = "Caleb Scott", Price = (decimal)25.77, Date= DateTime.Now.AddDays(-8)},
                new Order2{ Number = 10, Item = "Electric Steam Mop", Customer = "Caleb Scott", Price = (decimal)79.99, Date= DateTime.Now.AddDays(-9)},
                         
                new Order2{ Number = 1, Item = "Water Filter", Customer = "Morgan Wood", Price = (decimal)13.69, Date= DateTime.Now},
                new Order2{ Number = 2, Item = "Digital Scale", Customer = "Morgan Wood", Price = (decimal)13.00, Date= DateTime.Now.AddDays(-1) },
                new Order2{ Number = 3, Item = "Ceramic Heater", Customer = "Morgan Wood", Price = (decimal)24.97, Date= DateTime.Now.AddDays(-2)},
                new Order2{ Number = 4, Item = "Humidifier", Customer = "Morgan Wood", Price = (decimal)88.79, Date= DateTime.Now.AddDays(-3)},
                new Order2{ Number = 5, Item = "Coffee Filters", Customer = "Morgan Wood", Price = (decimal)7.79, Date= DateTime.Now.AddDays(-4)},
                new Order2{ Number = 6, Item = "Crock Pot", Customer = "Morgan Wood", Price = (decimal)23.99, Date= DateTime.Now.AddDays(-5)},
                new Order2{ Number = 7, Item = "Vacuum Cleaner", Customer = "Morgan Wood", Price = (decimal)79.00, Date= DateTime.Now.AddDays(-6)},
                new Order2{ Number = 8, Item = "Coffee Pot", Customer = "Morgan Wood", Price = (decimal)24.92, Date= DateTime.Now.AddDays(-7)},
                new Order2{ Number = 9, Item = "TV Stand", Customer = "Morgan Wood", Price = (decimal)25.77, Date= DateTime.Now.AddDays(-8)},
                new Order2{ Number = 10, Item = "Electric Steam Mop", Customer = "Morgan Wood", Price = (decimal)79.99, Date= DateTime.Now.AddDays(-9)},
                         
                new Order2{ Number = 1, Item = "Water Filter", Customer = "James Parker", Price = (decimal)13.69, Date= DateTime.Now},
                new Order2{ Number = 2, Item = "Digital Scale", Customer = "James Parker", Price = (decimal)13.00, Date= DateTime.Now.AddDays(-1) },
                new Order2{ Number = 3, Item = "Ceramic Heater", Customer = "James Parker", Price = (decimal)24.97, Date= DateTime.Now.AddDays(-2)},
                new Order2{ Number = 4, Item = "Humidifier", Customer = "James Parker", Price = (decimal)88.79, Date= DateTime.Now.AddDays(-3)},
                new Order2{ Number = 5, Item = "Coffee Filters", Customer = "James Parker", Price = (decimal)7.79, Date= DateTime.Now.AddDays(-4)},
                new Order2{ Number = 6, Item = "Crock Pot", Customer = "James Parker", Price = (decimal)23.99, Date= DateTime.Now.AddDays(-5)},
                new Order2{ Number = 7, Item = "Vacuum Cleaner", Customer = "James Parker", Price = (decimal)79.00, Date= DateTime.Now.AddDays(-6)},
                new Order2{ Number = 8, Item = "Coffee Pot", Customer = "James Parker", Price = (decimal)24.92, Date= DateTime.Now.AddDays(-7)},
                new Order2{ Number = 9, Item = "TV Stand", Customer = "James Parker", Price = (decimal)25.77, Date= DateTime.Now.AddDays(-8)},
                new Order2{ Number = 10, Item = "Electric Steam Mop", Customer = "James Parker", Price = (decimal)79.99, Date= DateTime.Now.AddDays(-9)},
                         
                new Order2{ Number = 1, Item = "Water Filter", Customer = "Austin Jackson", Price = (decimal)13.69, Date= DateTime.Now},
                new Order2{ Number = 2, Item = "Digital Scale", Customer = "Austin Jackson", Price = (decimal)13.00, Date= DateTime.Now.AddDays(-1) },
                new Order2{ Number = 3, Item = "Ceramic Heater", Customer = "Austin Jackson", Price = (decimal)24.97, Date= DateTime.Now.AddDays(-2)},
                new Order2{ Number = 4, Item = "Humidifier", Customer = "Austin Jackson", Price = (decimal)88.79, Date= DateTime.Now.AddDays(-3)},
                new Order2{ Number = 5, Item = "Coffee Filters", Customer = "Austin Jackson", Price = (decimal)7.79, Date= DateTime.Now.AddDays(-4)},
                new Order2{ Number = 6, Item = "Crock Pot", Customer = "Austin Jackson", Price = (decimal)23.99, Date= DateTime.Now.AddDays(-5)},
                new Order2{ Number = 7, Item = "Vacuum Cleaner", Customer = "Austin Jackson", Price = (decimal)79.00, Date= DateTime.Now.AddDays(-6)},
                new Order2{ Number = 8, Item = "Coffee Pot", Customer = "Austin Jackson", Price = (decimal)24.92, Date= DateTime.Now.AddDays(-7)},
                new Order2{ Number = 9, Item = "TV Stand", Customer = "Austin Jackson", Price = (decimal)25.77, Date= DateTime.Now.AddDays(-8)},
                new Order2{ Number = 10, Item = "Electric Steam Mop", Customer = "Austin Jackson", Price = (decimal)79.99, Date= DateTime.Now.AddDays(-9)},
            };
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