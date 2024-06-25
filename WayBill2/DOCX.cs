using OfficeOpenXml;
using Word = Microsoft.Office.Interop.Word;

namespace WayBill2
{
    public static class DOCX
    {
        public static void GetData(string excelFilePath)
        {
            // Открытие файла Excel
            ExcelPackage excelPackage = new ExcelPackage(excelFilePath);
            ExcelWorksheet dataTable = excelPackage.Workbook.Worksheets["Поставки товаров со склада"];
            
            List<Product> data = new List<Product>();

            for (int i = 1; i <= dataTable.Dimension.Rows; i++)
            {
                string productName = dataTable.Cells[i, 1].Text;
                int productQuantity = int.Parse(dataTable.Cells[i, 2].Text);
                int productPrice = int.Parse(dataTable.Cells[i,3].Text);
                string productProvider = dataTable.Cells[i, 4].Text;
                string productRecipient = dataTable.Cells[i, 5].Text;
                DateTime productDate = DateTime.Parse(dataTable.Cells[i, 6].Text);
                data.Add(new Product(productName, productQuantity, productPrice, productProvider, productRecipient, productDate));
            }

            foreach(Product product in data)
            {
                Console.WriteLine($"{product.Name} {product.Quantity} {product.Price} {product.SupplierName} {product.RecipientName} {product.Date}");
            }

            // Формирование накладных
            PrintData(data);

        }

        public static void PrintData(List<Product> data)
        {
            Dictionary<string, List<Product>> accordance = new Dictionary<string, List<Product>>();
            HashSet<DateTime> dates = new HashSet<DateTime>();

            foreach (Product product in data)
            {
                dates.Add(product.Date);
                if (accordance.ContainsKey(product.RecipientName)){
                    accordance[product.RecipientName].Add(product);
                }
                else
                {
                    accordance.Add(product.RecipientName, new List<Product>() { product });
                }
            }
            List<Product> temp = new List<Product>();
            foreach (string recipient in accordance.Keys)
            {
                foreach (DateTime date in dates)
                {
                    foreach (Product product in data)
                    {
                        if(product.Date == date && product.RecipientName == recipient)
                        {
                            temp.Add(product);
                        }
                    }
                    WriteWaybill(temp);
                    temp.Clear();
                }
            }

        }
        public static void WriteWaybill(List<Product> temp)
        {
            Random rand = new Random();
            Word.Application wordApp = new Word.Application();
            Word.Document doc = wordApp.Documents.Open($"{Directory.GetCurrentDirectory()}\\Товарная накладная.docx");

            try
            {
                string fileName = $"{Directory.GetCurrentDirectory()}\\Накладная от {temp[0].Date.Year}.{temp[0].Date.Month}.{temp[0].Date.Day}. Получатель - {temp[0].RecipientName}.docx";
                doc.SaveAs(fileName);
                doc.Close();

                doc = wordApp.Documents.Open(fileName);

                doc.Content.Find.Execute("<NUM>", ReplaceWith: $"{rand.Next(100000, 120000)}");
                doc.Content.Find.Execute("<DATE>", ReplaceWith: $"{temp[0].Date.ToShortDateString()}");
                doc.Content.Find.Execute("<PROVIDER>", ReplaceWith: $"{temp[0].SupplierName}");
                doc.Content.Find.Execute("<RECIPIENT>", ReplaceWith: $"{temp[0].RecipientName}");
                doc.Content.Find.Execute("<RECIPIENT>", ReplaceWith: $"{temp[0].RecipientName}");

                Word.Table table = doc.Tables[1];

                int summ = 0;

                // Заполнение ячеек таблицы
                for (int i = 1; i < temp.Count; i++)
                {
                    table.Rows.Add();
                    table.Cell(i + 1, 1).Range.Text = i.ToString();
                    table.Cell(i + 1, 2).Range.Text = temp[i].Name;
                    table.Cell(i + 1, 3).Range.Text = temp[i].Quantity.ToString();
                    table.Cell(i + 1, 4).Range.Text = temp[i].Price.ToString();

                    int total = (int)temp[i].Price * temp[i].Quantity;

                    table.Cell(i + 1, 5).Range.Text = total.ToString();

                    summ += total;
                }

                doc.Content.Find.Execute("<SUMM>", ReplaceWith: $"{summ}");
                doc.Content.Find.Execute("<SUMM>", ReplaceWith: $"{summ}");
                doc.Content.Find.Execute("<QUANTITY>", ReplaceWith: $"{temp.Count}");

                doc.SaveAs(fileName);
                Console.WriteLine($"Документ: {fileName} успешно создан");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
            finally
            {
                doc.Close();
                wordApp.Quit();
            }
        }

    }
}

        
