using System;
using System.Data.SqlClient;
using System.IO;
using OfficeOpenXml;

namespace ConsoleApp
{
    class Program
    {
        static void Main(string[] args)
        {
            // Set the LicenseContext for EPPlus
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            // Step 2: Get data from the user and save it to the SQL database
            Console.Write("Enter Name: ");
            string name = Console.ReadLine();

            Console.Write("Enter Roll: ");
            int roll = int.Parse(Console.ReadLine());

            Console.Write("Enter Subject: ");
            string subject = Console.ReadLine();

            SaveToDatabase(name, roll, subject);

            // Step 3: Write data to a text file
            WriteToTextFile(name, roll, subject);

            // Step 4: Write data to an Excel file
            WriteToExcelFile(name, roll, subject);

            // Step 5: Read files from the folder drive and show the file list in console output
            string folderPath = @"C:\Users\alif1\Downloads\TestConsoleApp\DataFile";
            ShowFileList(folderPath);

            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
        }

        // Method to save data to the SQL database
        private static void SaveToDatabase(string name, int roll, string subject)
        {
            string connectionString = "Server=NAHIYEAN-ALIF;Database=StudentTestDb;Trusted_Connection=true;";

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                string query = "INSERT INTO Students (Name, Roll, Subject) VALUES (@Name, @Roll, @Subject)";
                connection.Open();
                using (SqlCommand command = new SqlCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@Name", name);
                    command.Parameters.AddWithValue("@Roll", roll);
                    command.Parameters.AddWithValue("@Subject", subject);
                    command.ExecuteNonQuery();
                }
            }
        }


        private static void WriteToTextFile(string name, int roll, string subject)
        {
            string textFilePath = @"C:\Users\alif1\Downloads\TestConsoleApp\DataFile\StudentData.txt";
            string data = $"{name}, {roll}, {subject}";
            File.AppendAllText(textFilePath, data + Environment.NewLine);
        }

        // Method to write data to an Excel file and append new data to existing worksheet
        private static void WriteToExcelFile(string name, int roll, string subject)
        {
            string sheetName = "Sheet1"; // Default sheet name

            string excelFilePath = $@"C:\Users\alif1\Downloads\TestConsoleApp\DataFile\StudentData.xlsx";
            FileInfo file = new FileInfo(excelFilePath);

            using (ExcelPackage package = new ExcelPackage(file))
            {
                // Check if the sheet with the given name already exists
                ExcelWorksheet worksheet = package.Workbook.Worksheets[sheetName];

                if (worksheet == null)
                {
                    // If the sheet doesn't exist, add a new sheet with the given name
                    worksheet = package.Workbook.Worksheets.Add(sheetName);
                }

                // Find the last used row in the worksheet
                int lastUsedRow = worksheet.Dimension?.End.Row ?? 0;

                // Write data to the next empty row in the selected worksheet
                worksheet.Cells[lastUsedRow + 1, 1].Value = name;
                worksheet.Cells[lastUsedRow + 1, 2].Value = roll;
                worksheet.Cells[lastUsedRow + 1, 3].Value = subject;

                package.Save();
            }
        }

        // Method to show the list of files in the specified folder
        private static void ShowFileList(string folderPath)
        {
            string[] fileNames = Directory.GetFiles(folderPath);
            Console.WriteLine("List of Files:");
            foreach (string fileName in fileNames)
            {
                Console.WriteLine(fileName);
            }
        }
    }
}