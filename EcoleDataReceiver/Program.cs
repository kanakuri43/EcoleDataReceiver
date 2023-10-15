using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Xml.Linq;
using OpenPop.Mime;
using OpenPop.Pop3;

namespace EcoleDataReceiver
{
    class Program
    {
        static void Main()
        {
            try
            {
                Log("Starting the process...");

                var config = LoadConfig();

                var csvData = FetchEmailAttachment(config);

                InsertOrUpdateData(csvData, config);

                Log("Process completed successfully.");
            }
            catch (Exception ex)
            {
                Log($"Error: {ex.Message}");
            }
        }

        static dynamic LoadConfig()
        {
            var doc = XDocument.Load("config.xml");

            return new
            {
                ConnectionString = doc.Root.Element("Database").Element("ConnectionString").Value,
                Password = doc.Root.Element("Database").Element("Password").Value,
                TableName = doc.Root.Element("Database").Element("TableName").Value,
                PopServer = doc.Root.Element("Email").Element("PopServer").Value,
                PopPort = int.Parse(doc.Root.Element("Email").Element("PopPort").Value),
                EmailUser = doc.Root.Element("Email").Element("EmailUser").Value,
                EmailPassword = doc.Root.Element("Email").Element("EmailPassword").Value
            };
        }

        static DataTable FetchEmailAttachment(dynamic config)
        {
            Log("Fetching email attachment...");

            using var client = new Pop3Client();
            client.Connect(config.Pop3Server, config.Pop3Port, true);
            client.Authenticate(config.Pop3User, config.Pop3Password);

            var messageCount = client.GetMessageCount();
            var messages = new List<Message>();
            var attachmentMessage = messages.LastOrDefault(m => m.FindAllAttachments().Any());
            if (attachmentMessage == null)
            {
                throw new Exception("No email with attachments found.");
            }

            var csvAttachment = attachmentMessage.FindAllAttachments().First();
            var csvStream = new MemoryStream(csvAttachment.Body);

            var dataTable = new DataTable();
            // Parse the CSV data from the stream into the DataTable...
            // For this example, we assume simple comma-separated values
            using var reader = new StreamReader(csvStream);
            var headers = reader.ReadLine().Split(',');
            foreach (var header in headers)
            {
                dataTable.Columns.Add(header);
            }
            while (!reader.EndOfStream)
            {
                var line = reader.ReadLine();
                var values = line.Split(',');
                dataTable.Rows.Add(values);
            }

            Log("Fetched and parsed CSV attachment.");
            return dataTable;
        }

        static DataTable LoadCsvToDataTable(string filePath)
        {
            Log("Loading CSV to DataTable...");

            var dataTable = new DataTable();
            using var reader = new StreamReader(filePath);

            var headers = reader.ReadLine().Split(',');
            foreach (var header in headers)
            {
                dataTable.Columns.Add(header);
            }

            while (!reader.EndOfStream)
            {
                var line = reader.ReadLine();
                var values = line.Split(',');
                dataTable.Rows.Add(values);
            }

            return dataTable;
        }

        static void InsertOrUpdateData(DataTable dataTable, dynamic config)
        {
            Log("Inserting or updating data...");

            using var connection = new SqlConnection(config.ConnectionString);
            connection.Open();

            // This is just a simple example. In a real scenario, 
            // the logic for inserting or updating would be more complex.
            foreach (DataRow row in dataTable.Rows)
            {
                var cmdText = $"INSERT INTO your_table_name (col1, col2, ... ) VALUES (@value1, @value2, ... )";
                // Alternatively, use an UPDATE statement if required.

                using var command = new SqlCommand(cmdText, connection);
                command.Parameters.AddWithValue("@value1", row["col1"]);
                command.Parameters.AddWithValue("@value2", row["col2"]);
                // ... Add other parameters ...

                command.ExecuteNonQuery();
            }

            Log("Data inserted/updated successfully.");
        }


        static void Log(string message)
        {
            var logFolder = "./log/";
            if (!Directory.Exists(logFolder))
            {
                Directory.CreateDirectory(logFolder);
            }

            var logPath = $"{logFolder}{DateTime.Now:yyyyMMdd-HHmmss}.log";
            using var writer = new StreamWriter(logPath, true);
            writer.WriteLine($"[{DateTime.Now}] {message}");
        }
    }
}