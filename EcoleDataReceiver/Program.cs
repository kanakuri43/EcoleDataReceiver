using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Net.Security;
using System.Text;
using System.Xml.Linq;
using EcoleDataReceiver.Models;
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
                //Console.WriteLine(string.Format("{0} Starting the process...", DateTime.Now.ToString("HH:mm:ss")));
                Console.WriteLine($"{DateTime.Now.ToString("HH:mm:ss")} Starting the process...");

                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

                dynamic config = LoadConfig();

                string csvFilePath = ReceiveMailAndSaveAttachment(config);
                if (csvFilePath == "") 
                {
                    Console.WriteLine($"{DateTime.Now.ToString("HH:mm:ss")} No email with subject containing '{config.subjectContains}' and attachments found.");
                    return;
                }

                DataTable dataTable = LoadCSV(csvFilePath);

                InsertOrUpdateData(dataTable, config);

                Console.WriteLine(string.Format("{0} Process completed successfully.", DateTime.Now.ToString("HH:mm:ss")));
            }
            catch (Exception ex)
            {
                //Console.WriteLine(string.Format("{0} Error: {1}", DateTime.Now.ToString("HH:mm:ss"), ex.Message));
                Console.WriteLine($"{DateTime.Now.ToString("HH:mm:ss")} {ex.Message}");

            }
            finally
            {
                sw.Dispose();
            }
        }

        static dynamic LoadConfig()
        {
            var doc = XDocument.Load("config.xml");

            return new
            {
                ConnectionString = doc.Root.Element("Database").Element("ConnectionString").Value,
                InputFolder = doc.Root.Element("Input").Element("Folder").Value,
                Pop3Server = doc.Root.Element("Email").Element("Pop3Server").Value,
                Pop3Port = int.Parse(doc.Root.Element("Email").Element("Pop3Port").Value),
                Pop3User = doc.Root.Element("Email").Element("Pop3User").Value,
                Pop3Password = doc.Root.Element("Email").Element("Pop3Password").Value,
                IdentifySubject = doc.Root.Element("Email").Element("IdentifySubject").Value
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
                if (message.Headers.Subject.Contains(config.IdentifySubject) && message.FindAllAttachments().Any())
                {
                    attachmentMessage = message;
                    break;
                }
            }

            if (attachmentMessage == null)
            {
                //throw new Exception($"No email with subject containing '{config.subjectContains}' and attachments found.");
                return "";
            }

            // Get the first CSV attachment
            var csvAttachment = attachmentMessage.FindAllAttachments().FirstOrDefault(att => att.FileName.EndsWith(".csv"));
            if (csvAttachment == null)
            {
                throw new Exception("No CSV attachments found.");
            }

            // Save the CSV attachment to a file
            var filePath = Path.Combine(Environment.CurrentDirectory, $"{config.InputFolder}{csvAttachment.FileName}");
            File.WriteAllBytes(filePath, csvAttachment.Body);

            Console.WriteLine($"{DateTime.Now:HH:mm:ss} Attachment saved to {filePath}");
            return filePath;
        }



        static DataTable LoadCSV(string filePath)
        {
            Console.WriteLine($"{DateTime.Now:HH:mm:ss} Loading CSV File...");

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
            Console.WriteLine($"{DateTime.Now:HH:mm:ss} Inserting or updating data...");
            /*

            using var connection = new SqlConnection(config.ConnectionString);
            connection.Open();

            // This is just a simple example. In a real scenario, 
            // the logic for inserting or updating would be more complex.
            foreach (DataRow row in dataTable.Rows)
            {
                var sql = $"INSERT INTO your_table_name (col1, col2, ... ) VALUES (@value1, @value2, ... )";
                // Alternatively, use an UPDATE statement if required.

                using var command = new SqlCommand(sql, connection);
                command.Parameters.AddWithValue("@value1", row["col1"]);
                command.Parameters.AddWithValue("@value2", row["col2"]);
                // ... Add other parameters ...

                command.ExecuteNonQuery();
            }
            */
            foreach (DataRow row in dataTable.Rows)
            {
                using (var context = new AppDbContext())
                {

                    var product = context.Products.FirstOrDefault(p => p.Id == (int)row[0] && p.State == 0);
                    if (product != null)
                    {
                        // Update

                        product.ProductName = row[1].ToString();    // 商品名
                        //product.State = (int)row[2];
                        //product.Sundry = (int)row[0];
                        //product.TaxationType = (int)row[0];
                        product.ProductCategoryId = (int)row[3];    // 分類コード
                        product.Unit = row[4].ToString();           // 単位
                        //product.Price = (int)row[0];
                        product.Cost = (int)row[0];
                        product.CatalogPrice = (int)row[0];
                        product.StockType = (int)row[0];
                        product.JAN = (int)row[0];
                        product.ReserveNo1 = (int)row[0];
                        product.ReserveNo2 = (int)row[0];
                        product.ProdutNo = row[2].ToString();
                        product.MakerName = row[1].ToString();
                        product.IsReadOnly = (int)row[0];
                        product.UpdatedAt = (int)row[0];

                        context.SaveChanges();
                    }
                }
            }

            Console.WriteLine($"{DateTime.Now:HH:mm:ss} Data inserted/updated successfully.");
        }

    }
}