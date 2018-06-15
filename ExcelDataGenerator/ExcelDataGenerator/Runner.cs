namespace ExcelDataGenerator
{
    public class Runner
    {
        public void Run()
        {
            var options = new ExcelDataGeneratorOptions()
            {
                StartDate = new System.DateTime(2017, 01, 01),
                EndDate = new System.DateTime(2017, 12, 31),
                Customers = new System.Collections.Generic.List<Customer>()
                {
                    new Customer() { Id = 1, Name = "Klient A", Address = "Adres A", City = "Miasto A", Country = "Polska", PostCode = "00-000"},
                    new Customer() { Id = 2, Name = "Klient B", Address = "Adres B", City = "Miasto B", Country = "Polska", PostCode = "00-000"},
                    new Customer() { Id = 3, Name = "Klient C", Address = "Adres C", City = "Miasto C", Country = "Polska", PostCode = "00-000"},
                    new Customer() { Id = 4, Name = "Klient D", Address = "Adres D", City = "Miasto D", Country = "Polska", PostCode = "00-000"},
                    new Customer() { Id = 5, Name = "Klient E", Address = "Adres E", City = "Miasto E", Country = "Polska", PostCode = "00-000"},
                    new Customer() { Id = 6, Name = "Klient F", Address = "Adres F", City = "Miasto F", Country = "Polska", PostCode = "00-000"},
                    new Customer() { Id = 7, Name = "Klient G", Address = "Adres G", City = "Miasto G", Country = "Polska", PostCode = "00-000"},
                    new Customer() { Id = 8, Name = "Klient H", Address = "Adres H", City = "Miasto H", Country = "Polska", PostCode = "00-000"}
                }
            };
        }
    }
}
