using System;
using OfficeOpenXml;
using System.Data.SqlClient;
using DotNetEnv;


namespace ExcelReader
{
    internal class Program
    {
        //Filepath to excel
        public static string filePath = "C:\\Users\\Milana\\Documents\\Test.xlsx";

        public static string connectionString = "";

        static void Main(string[] args)
        {
            // Set the license context
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            // Load connection string out of ENV
            string envFilePath = @"C:\Users\Milana\Documents\ExcelReader\.env";
            DotNetEnv.Env.Load(envFilePath);
            connectionString = Environment.GetEnvironmentVariable("CONNECTION_STRING");


            if(connectionString != null ) 
            {
                bool tableExists = CheckIfTableExists();

                if (tableExists)
                {
                    //DeleteTable();
                }

                string[] columnList = GetExcelColumnNames(filePath);

                CreateTable(columnList);
            }
        }

        static bool CheckIfTableExists()
        {
            bool tableExists = false;

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();

                    string query = $"SELECT COUNT(*) FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'ExcelData'";

                    using (SqlCommand command = new SqlCommand(query, connection))
                    {
                        int count = Convert.ToInt32(command.ExecuteScalar());
                        tableExists = (count > 0);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }

            return tableExists;
        }

        static void DeleteTable()
        {
            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();

                    string query = $"DROP TABLE IF EXISTS ExcelData";

                    using (SqlCommand command = new SqlCommand(query, connection))
                    {
                        command.ExecuteNonQuery();
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }

        static void CreateTable (string[] columnNames)
        {
            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();

                    // Query to create the table with columns and their datatype
                    string createQuery = $"CREATE TABLE ExcelData (" +
                                         string.Join(", ", columnNames.Select(col => $"{col} NVARCHAR(MAX)")) + 
                                         ")";

                    using (SqlCommand command = new SqlCommand(createQuery, connection))
                    {
                        command.ExecuteNonQuery();
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }

        static string[] GetExcelColumnNames(string filePath)
        {
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                // Get the first worksheet only
                var worksheet = package.Workbook.Worksheets.First();

                // Get the first row
                var firstRow = worksheet.Cells["1:1"];

                // Extract column names from the first row
                string[] columnNames = firstRow.Select(cell => cell.Text).ToArray();

                return columnNames;
            }
        }
    }
}