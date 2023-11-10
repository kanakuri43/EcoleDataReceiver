using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Net.Security;
using System.Xml.Linq;
using OpenPop.Mime;
using OpenPop.Pop3;

namespace EcoleDataReceiver
{
    class Program
    {
        static void Main()
        {
            StreamWriter sw = new StreamWriter(string.Format(@"log/{0}.log", DateTime.Now.ToString("yyyyMMdd_HHmmss"), true));
            try
            {
                Console.SetOut(sw); // 出力先を設定
                Console.WriteLine(string.Format("{0} Starting the process...", DateTime.Now.ToString("HH:mm:ss")));

                dynamic config = LoadConfig();

                string csvFilePath = ReceiveMailAndSaveAttachment(config);

                DataTable dataTable = LoadCSV(csvFilePath);

                InsertOrUpdateData(dataTable, config);

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
                Pop3Server = doc.Root.Element("Email").Element("Pop3Server").Value,
                Pop3Port = int.Parse(doc.Root.Element("Email").Element("Pop3Port").Value),
                Pop3User = doc.Root.Element("Email").Element("Pop3User").Value,
                Pop3Password = doc.Root.Element("Email").Element("Pop3Password").Value,
                subjectContains = "asdf"
            };
        }

        static string ReceiveMailAndSaveAttachment(dynamic config)
        {
            Console.WriteLine($"{DateTime.Now:HH:mm:ss} Fetching email attachment...");

            using var client = new Pop3Client();
            client.Connect(config.Pop3Server, config.Pop3Port, true);
            client.Authenticate(config.Pop3User, config.Pop3Password);

            var messageCount = client.GetMessageCount();

            // Find the first message with the specified subject and an attachment
            Message attachmentMessage = null;
            for (int i = messageCount; i >= 1; i--)
            {
                var message = client.GetMessage(i);
                if (message.Headers.Subject.Contains(config.subjectContains) && message.FindAllAttachments().Any())
                {
                    attachmentMessage = message;
                    break;
                }
            }

            if (attachmentMessage == null)
            {
                throw new Exception($"No email with subject containing '{config.subjectContains}' and attachments found.");
            }

            // Get the first CSV attachment
            var csvAttachment = attachmentMessage.FindAllAttachments().FirstOrDefault(att => att.FileName.EndsWith(".csv"));
            if (csvAttachment == null)
            {
                throw new Exception("No CSV attachments found.");
            }

            // Save the CSV attachment to a file
            var filePath = Path.Combine(Environment.CurrentDirectory, csvAttachment.FileName);
            File.WriteAllBytes(filePath, csvAttachment.Body);

            Console.WriteLine($"Attachment saved to {filePath}");
            return filePath;
        }



        static DataTable LoadCSV(string filePath)
        {
            DataTable dataTable = new DataTable();

            // CSVファイルを読み込む
            using (StreamReader sr = new StreamReader(filePath))
            {
                // ヘッダー行を読み込んでDataTableのカラムを設定
                string[] headers = sr.ReadLine().Split(',');
                foreach (string header in headers)
                {
                    dataTable.Columns.Add(header);
                }

                // CSVの各行を読み込んでDataTableの行に追加
                while (!sr.EndOfStream)
                {
                    string[] rows = sr.ReadLine().Split(',');
                    DataRow dr = dataTable.NewRow();
                    for (int i = 0; i < headers.Length; i++)
                    {
                        dr[i] = rows[i];
                    }
                    dataTable.Rows.Add(dr);
                }
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