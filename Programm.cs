using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using Dapper;
using Z.Dapper.Plus;
using ExcelDataReader;

class Program
{
    static void Main(string[] args)
    {
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        var salesList = new List<Sale>();

        using (var stream = File.Open("sales_data.xlsx", FileMode.Open, FileAccess.Read))
        {
            using (var reader = ExcelReaderFactory.CreateReader(stream))
            {
                var result = reader.AsDataSet();
                var table = result.Tables[0];

                for (int i = 1; i < table.Rows.Count; i++)
                {
                    salesList.Add(new Sale
                    {
                        Id = Convert.ToInt32(table.Rows[i][0]),
                        ProductName = table.Rows[i][1].ToString(),
                        Category = table.Rows[i][2].ToString(),
                        Price = Convert.ToDecimal(table.Rows[i][3]),
                        Quantity = Convert.ToInt32(table.Rows[i][4]),
                        SaleDate = Convert.ToDateTime(table.Rows[i][5])
                    });
                }
            }
        }

        string connectionString = "Server=.;Database=SalesDb;Trusted_Connection=True;";

        DapperPlusManager.Entity<Sale>().Table("Sales");

        using (var connection = new SqlConnection(connectionString))
        {
            connection.Open();
            connection.BulkInsert(salesList);
        }

        Console.WriteLine("Data imported successfully!");
    }
}

public class Sale
{
    public int Id { get; set; }
    public string ProductName { get; set; }
    public string Category { get; set; }
    public decimal Price { get; set; }
    public int Quantity { get; set; }
    public DateTime SaleDate { get; set; }
}
