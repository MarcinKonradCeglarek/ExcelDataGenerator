using System;
using System.Collections.Generic;

namespace ExcelDataGenerator
{
    public class ExcelDataGenerator
    {
        static Random rnd = new Random();

        private ExcelDataGeneratorOptions options;

        public ExcelDataGenerator(ExcelDataGeneratorOptions options)
        {
            this.options = options;
        }

        public List<Order> GenerateOrders()
        {
            var orderId = 1000 + rnd.Next(1000);
            var retVal = new List<Order>();
            for (var currentDate = this.options.StartDate; currentDate <= this.options.EndDate; currentDate = currentDate.AddDays(1))
            {
                retVal.Add(this.GenerateOrder(orderId++, currentDate));
                retVal.Add(this.GenerateOrder(orderId++, currentDate));
                retVal.Add(this.GenerateOrder(orderId++, currentDate));
            }

            return retVal;
        }

        private Order GenerateOrder(int id, DateTime date)
        {
            var customer = this.options.Customers[rnd.Next(this.options.Customers.Count)];
            var salesPerson = this.options.SalesPeople[rnd.Next(this.options.SalesPeople.Count)];
            var product = this.options.Products[rnd.Next(this.options.Products.Count)];
            var productCount = rnd.Next(20);

            var retVal = new Order()
            {
                OrderId = id,
                OrderDate = date,
                CustomerId = customer.Id,
                CustomerName = customer.Name,
                CustomerAddress = customer.Address,
                CustomerCity = customer.City,
                CustomerPostCode = customer.PostCode,
                CustomerCountry = customer.Country,
                SalesPersonName = salesPerson.Name,
                SalesPersonRegion = salesPerson.Region,
                ProductName = product.Name,
                ProductCategory = product.Category,
                ProductQuantity = productCount,
                ProductUnitPrice = product.UnitPrice,
                ShippingCost = (decimal)Math.Sqrt(productCount) * product.ShippingPrice
            };

            return retVal;
        }
    }

    public class ExcelDataGeneratorOptions
    {
        public DateTime StartDate { get; set; }

        public DateTime EndDate { get; set; }

        public List<Product> Products { get; set; }

        public List<Customer> Customers { get; set; }

        public List<SalesPerson> SalesPeople { get; set; }
    }

    public class Product
    {
        public string Name { get; set; }

        public string Category { get; set; }

        public decimal UnitPrice { get; set; }

        public decimal ShippingPrice { get; set; }
    }

    public class Customer
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public string Address { get; set; }
        public string City { get; set; }
        public string PostCode { get; set; }
        public string Country { get; set; }
    }

    public class SalesPerson
    {
        public string Name { get; set; }
        public string Region { get; set; }
    }

    public class Order
    {
        public int OrderId { get; set; }
        public DateTime OrderDate { get; set; }
        public int Quarter => ((int)((this.OrderDate.Month - 1) / 3)) + 1;
        public string Month => this.OrderDate.ToString("MMMM");
        public int CustomerId { get; set; }
        public string CustomerName { get; set; }
        public string CustomerAddress { get; set; }
        public string CustomerCity { get; set; }
        public string CustomerPostCode { get; set; }
        public string CustomerCountry { get; set; }
        public string SalesPersonName { get; set; }
        public string SalesPersonRegion { get; set; }
        public string ProductName { get; set; }
        public string ProductCategory { get; set; }
        public decimal ProductUnitPrice { get; set; }
        public int ProductQuantity { get; set; }
        public decimal ProductTotalPrice => this.ProductUnitPrice * this.ProductQuantity;
        public decimal ShippingCost { get; set; }
    }
}
