using CsvHelper.Configuration;
using CsvHelper;
using System.Globalization;

namespace WayBill2
{
    public class CSV
    {
        public static List<Product> ReadCsv(string filePath)
        {
            StreamReader reader = new StreamReader(filePath);
            CsvReader csv = new CsvReader(reader, new CsvConfiguration(CultureInfo.InvariantCulture));
            
               List<Product> forRecords = csv.GetRecords<Product>().ToList();

            reader.Close();

            List<Product> temporaryProduct = new List<Product>();

            foreach (Product product in forRecords)
            {
                bool isDataExists = false;
                for (int i = 0; i < temporaryProduct.Count; i++)
                {
                    if (temporaryProduct[i].Name == product.Name && temporaryProduct[i].RecipientName == product.RecipientName && temporaryProduct[i].Date == product.Date)
                    {
                        temporaryProduct[i].Quantity += product.Quantity;
                        isDataExists = true;
                    }
                }

                if (!isDataExists)
                {
                    temporaryProduct.Add(product);
                }
            }
            temporaryProduct.Sort();
            return temporaryProduct;
        }
    }
}
