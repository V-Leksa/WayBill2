using OfficeOpenXml;

namespace WayBill2
{
    public static class XLSX
    {
        public static void CreateExcelFile(List<Product> products, string filePath)
        {
            ExcelPackage package = new ExcelPackage(filePath);
            ExcelWorksheet dataTable;

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            bool isSheetExists = false;
            foreach (ExcelWorksheet worksheet in package.Workbook.Worksheets)
            {
                if(worksheet.Name == "Поставки товаров со склада")
                {
                    isSheetExists = true;
                }
            }
            if(isSheetExists)
            {
                dataTable = package.Workbook.Worksheets["Поставки товаров со склада"];
            }
            else
            {
                dataTable = package.Workbook.Worksheets.Add("Поставки товаров со склада");
            }

            for (int i = 0; i < products.Count; i++)
            {
                dataTable.Cells[i + 1, 1].Value = products[i].Name;
                dataTable.Cells[i + 1, 2].Value = products[i].Quantity;
                dataTable.Cells[i + 1, 3].Value = products[i].Price;
                dataTable.Cells[i + 1, 4].Value = products[i].SupplierName;
                dataTable.Cells[i + 1, 5].Value = products[i].RecipientName;
                dataTable.Cells[i + 1, 6].Value = products[i].Date.Date;
            }
            package.Save();
        }
    }
}
