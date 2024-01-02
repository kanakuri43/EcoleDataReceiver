using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Data.SQLite;
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
using static Microsoft.EntityFrameworkCore.DbLoggerCategory;

namespace EcoleDataReceiver
{
    class Program
    {
        static void Main()
        {
            StreamWriter sw = new StreamWriter(string.Format(@"log/{0}.log", DateTime.Now.ToString("yyyyMMdd_HHmmss"), true));
            try
            {
                // 初期設定
                Console.SetOut(sw); 
                Console.WriteLine($"{DateTime.Now.ToString("HH:mm:ss")} Starting the process...");
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                dynamic config = LoadConfig();

                string sqliteFileName = "";
                string receivedMailId = "";

                // inputフォルダが空かチェック
                if (InputDirectoryEmptyCheck(config) == true)
                {
                    // 空ならばメール受信
                    //ReceiveMailAndSaveAttachment(config, ref sqliteFileName, ref receivedMailId);
                }

                // 受信したTSVからDatatableを作る
                //DataTable dataTable = CreateDatatableFromTsv(config);
                DataTable dataTable = CreateDatatableFromSQLite(config);

                // DB更新
                InsertOrUpdateSqlServer(config, dataTable);

                // 更新済みのsqliteファイル削除
                DeleteUpdatedSqliteFiles(config, sqliteFileName);

                // 更新済みのメール削除

                Console.WriteLine(string.Format("{0} Process completed successfully.", DateTime.Now.ToString("HH:mm:ss")));
            }
            catch (Exception ex)
            {
                Console.WriteLine($"{DateTime.Now.ToString("HH:mm:ss")} Error: {ex.Message}");

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

                Pop3Server = doc.Root.Element("Email").Element("Pop3").Element("Server").Value,
                Pop3Port = int.Parse(doc.Root.Element("Email").Element("Pop3").Element("Port").Value),
                Pop3User = doc.Root.Element("Email").Element("Pop3").Element("User").Value,
                Pop3Password = doc.Root.Element("Email").Element("Pop3").Element("Password").Value,

                EmailTo = doc.Root.Element("Email").Element("Smtp").Element("To").Value,
                EmailFrom = doc.Root.Element("Email").Element("Smtp").Element("From").Value,
                SmtpServer = doc.Root.Element("Email").Element("Smtp").Element("Server").Value,
                SmtpPort = int.Parse(doc.Root.Element("Email").Element("Smtp").Element("Port").Value),
                SmtpUser = doc.Root.Element("Email").Element("Smtp").Element("User").Value,
                SmtpPassword = doc.Root.Element("Email").Element("Smtp").Element("Password").Value,

                Subject = doc.Root.Element("Email").Element("Smtp").Element("Subject").Value + doc.Root.Element("Company").Element("Id").Value
            };
        
        }

        static bool InputDirectoryEmptyCheck(dynamic config)
        {
            Console.WriteLine(string.Format("{0} Checking the input directory is empty...", DateTime.Now.ToString("HH:mm:ss")));

            return true;
        }


        static string ReceiveMailAndSaveAttachment(dynamic config, ref string filePath, ref string receivedMailId)
        {

            using var client = new Pop3Client();
            client.Connect(config.Pop3Server, config.Pop3Port, true);
            client.Authenticate(config.Pop3User, config.Pop3Password);

            var messageCount = client.GetMessageCount();

            Message attachmentMessage = null;
            for (int i = messageCount; i >= 1; i--)
            {
                var message = client.GetMessage(i);
                if (message.Headers.Subject.Contains(config.Subject) && message.FindAllAttachments().Any())
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
            var csvAttachment = attachmentMessage.FindAllAttachments().FirstOrDefault(att => att.FileName.EndsWith(".tsv"));
            if (csvAttachment == null)
            {
                throw new Exception("No attachments found.");
            }

            // Save the sqlite attachment to a file
            filePath = Path.Combine(Environment.CurrentDirectory, $"{config.InputFolder}{csvAttachment.FileName}");
            File.WriteAllBytes(filePath, csvAttachment.Body);

            Console.WriteLine($"{DateTime.Now:HH:mm:ss} Attachment saved to {filePath}");
            return filePath;
        }



        static DataTable CreateDatatableFromTsv(dynamic config)
        {
            Console.WriteLine($"{DateTime.Now:HH:mm:ss} Loading TSV File...");

            string filePath = "";
            string[] names = Directory.GetFiles($"{config.InputFolder}", "*.tsv");
            foreach (string name in names)
            {
                filePath = $"{config.InputFolder}" + name;
            }

            DataTable dataTable = new DataTable();

            using (StreamReader sr = new StreamReader(filePath))
            {
                // ヘッダー行を読み込んでDataTableのカラムを設定
                string[] headers = sr.ReadLine().Split('\t');
                foreach (string header in headers)
                {
                    dataTable.Columns.Add(header);
                }

                // CSVの各行を読み込んでDataTableの行に追加
                while (!sr.EndOfStream)
                {
                    string[] rows = sr.ReadLine().Split('\t');
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
        public static DataTable CreateDatatableFromSQLite(dynamic config)
        {

            string filePath = "";
            string[] names = Directory.GetFiles($"{config.InputFolder}", "*.sqlite");
            foreach (string name in names)
            {
                filePath = name;
            }


            DataTable dataTable = new DataTable();

            using (SQLiteConnection sqliteCon = new SQLiteConnection($"Data Source={filePath};Version=3;"))
            {
                sqliteCon.Open();

                string query = $"SELECT * FROM updated_items";
                using (SQLiteCommand command = new SQLiteCommand(query, sqliteCon))
                {
                    using (SQLiteDataAdapter adapter = new SQLiteDataAdapter(command))
                    {
                        adapter.Fill(dataTable);
                    }
                }
            }

            return dataTable;
        }
        static void InsertOrUpdateSqlServer(dynamic config, DataTable dataTable)
        {
            Console.WriteLine($"{DateTime.Now:HH:mm:ss} Insert or update data...");

            foreach (DataRow row in dataTable.Rows)
            {
                using (var context = new AppDbContext())
                {

                    var product = context.Products.FirstOrDefault(p => p.Id == (int)row[0] && p.State == 0);
                    if (product != null)
                    {
                        // Update

                        product.ProductName = row[1].ToString();    // 商品名
                        product.ProductCategoryId = (int)row[3];    // 分類コード
                        product.Unit = row[4].ToString();           // 単位
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
                    else
                    {
                        // Insert
                        var p = new Product
                        {
                            Id = (int)row[0],
                            State = (int)row[1],
                            Sundry = (int)row[2],
                            TaxationType = (int)row[3],
                            ProductCategoryId = (int)row[4],
                            Unit = row[5].ToString(),
                            Price = (int)row[6],
                            Cost = (int)row[7],
                            CatalogPrice = (int)row[8],
                            StockType = (int)row[9],
                            JAN = (int)row[10],
                            ReserveNo1 = (int)row[11],
                            ReserveNo2 = (int)row[12],
                            ProdutNo = row[13].ToString(),
                            MakerName = row[14].ToString(),
                            IsReadOnly = (int)row[15],
                            UpdatedAt = (int)row[16],
                        };
                        context.Products.Add(product);
                        context.SaveChanges();
                    }
                }
            }

            Console.WriteLine($"{DateTime.Now:HH:mm:ss} Data inserted/updated successfully.");
        }


        static void DeleteUpdatedSqliteFiles(dynamic config, string fileName)
        {

        }
    }
}