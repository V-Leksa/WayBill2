namespace WayBill2
{
     class Program
     {
        private static string _dataPath = "data.csv";
        static void Main(string[] args)
        {
            List<Product> product = CSV.ReadCsv(_dataPath);

            XLSX.CreateExcelFile(product, "output.xlsx");

            DOCX.GetData("output.xlsx");
        }
    
     }
}