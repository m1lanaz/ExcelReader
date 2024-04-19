using System;
using OfficeOpenXml;
using System.Data.SqlClient;
using DotNetEnv;


namespace ExcelReader
{
    internal class Program
    {
        //Filepath to excel
        public static string filePath = "";

        public static string connectionString = "";

        static void Main(string[] args)
        {
            // Load connection string out of ENV
            string envFilePath = @"C:\Users\Milana\Documents\ExcelReader\.env";
            DotNetEnv.Env.Load(envFilePath);
            connectionString = Environment.GetEnvironmentVariable("CONNECTION_STRING");

            if(connectionString != null ) 
            {
                bool tableExists = CheckIfTableExists();

                if (tableExists)
                {
                    DeleteTable();
                }



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

                    string query = $"DROP TABLE IF EXISTS EXCELDATA";

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
    }
}