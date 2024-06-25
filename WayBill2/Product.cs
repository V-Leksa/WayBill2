using CsvHelper.Configuration.Attributes;

namespace WayBill2
{
    public class Product : IComparable<Product>
    {
        // Атрибуты для соответствующих полей в CSV файле
        [Name("Название товара")]
        public string? Name { get; set; }

        [Name("Количество")]
        public int Quantity { get; set; }

        [Name("Стоимость")]
        public int Price { get; set; }

        [Name("ФИО поставщика")]
        public string? SupplierName { get; set; }

        [Name("ФИО получателя")]
        public string? RecipientName { get; set; }

        [Name("Дата поставки")]
        public DateTime Date { get; set; }

        public Product() { }

        public Product(string? name, int quantity, int price, string? supplierName, string? recipientName, DateTime date)
        {
            Name = name;
            Quantity = quantity;
            Price = price;
            SupplierName = supplierName;
            RecipientName = recipientName;
            Date = date;
        }

        // Реализация интерфейса IComparable для сортировки по дате поставки
        public int CompareTo(Product? obj)
        {
            return Date.CompareTo(obj?.Date);
        }
    }
}
