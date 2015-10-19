using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using MongoDB.Bson;
using MongoDB.Driver;
using System.Data.OleDb;
using System.Data;
using System.Security.Permissions;


namespace ImportingReports
{
    class Program
    {
        static void Main()
        {
            //var extractPath = ExtractingFromZip();
            var xmlPath =
                @"C:\Users\Cookie\Desktop\Telerik Academy\Databases\Database-2015-Team-Flerovium-\ImportingReports\ImportingReports\XmlImportFile\Books.xml";
            XDocument doc = XDocument.Load(xmlPath);

            var zipFilePath = @"C:\Users\Cookie\Desktop\Telerik Academy\Databases\Database-2015-Team-Flerovium-\ImportingReports\ImportingReports\ZipFile\Sales-Reports.zip";
            var extractedFolderPath = ExtractZip(zipFilePath);
            var dateNamedFolders = GetDateNamedFolders(extractedFolderPath);

            TestExcel(dateNamedFolders.ElementAt(0).GetFiles()[0].FullName, dateNamedFolders.ElementAt(0).GetFiles()[0].Name);

            var root = doc.Root;

            var books = root.Elements().AsQueryable().Select(BookMongo.FromXElement).ToList();

            var client = new MongoClient("mongodb://localhost");
            var db = client.GetDatabase("books-db");

            var booksCollection = db.GetCollection<BsonDocument>("books");

            var bsonBooks = books.AsQueryable().Select(BookMongo.ToBsonDocument);

            //booksCollection.InsertManyAsync(bsonBooks);
            //Console.ReadKey();


            /* In order to create a PDF report, use this: 
            var db = new dbEntities();   // or whatever we use      
            var testPdf = new PDFReporter(db); // creating an instance
            var data = db.Sales.ToList(); // or something more precise
            testPdf.GeneratePdfSalesReport("Sales", data); 
            */
        }

        static string ExtractZip(string zipFilePath)
        {
            var extractedZipFolder = Path.Combine(Path.GetTempPath(), "reports");
            if (Directory.Exists(extractedZipFolder))
            {
                Directory.Delete(extractedZipFolder, true);
            }
            ZipFile.ExtractToDirectory(zipFilePath, extractedZipFolder);
            return extractedZipFolder;
        }

        private static IEnumerable<DirectoryInfo> GetDateNamedFolders(string extractedFolderPath)
        {
            List<DirectoryInfo> dirs = new List<DirectoryInfo>();
            var dirInfo = new DirectoryInfo(extractedFolderPath);
            TraverseDirectory(dirInfo, dirs);
            return dirs;
        }

        private static void TraverseDirectory(DirectoryInfo dirInfo, IList<DirectoryInfo> dirs)
        {
            if (IsDate(dirInfo.Name))
            {
                dirs.Add(dirInfo);
                //Console.WriteLine(dirInfo.Name);
            }
            else
            {
                foreach (var dir in dirInfo.GetDirectories())
                {
                    TraverseDirectory(dir, dirs);
                }
            }
        }

        private static bool IsDate(string str)
        {
            try
            {
                DateTime.ParseExact(str, "dd-MMM-yyyy", CultureInfo.InvariantCulture);
                return true;
            }
            catch
            {
                return false;
            }
        }

        


        private static void TestExcel(string excelFilePath, string excelFileName)
        {
            var connectionStringFormat = @"Provider=Microsoft.Jet.OLEDB.4.0;
               Data Source={0};
               Extended Properties=Excel 8.0";

            var connectionString = string.Format(connectionStringFormat, excelFilePath);
            var conn = new OleDbConnection(connectionString);
            conn.Open();
            var command = new OleDbCommand("Select * From [Sales$]", conn);

            var adapter = new OleDbDataAdapter();
            adapter.SelectCommand = command;

            var dataSet = new DataSet();


            adapter.Fill(dataSet, excelFileName);
            var table = dataSet.Tables[0];
           
            Console.WriteLine(table.TableName);

            foreach (DataRow data in table.Rows)
            {
                // var product = new Products();
                //product.Id = data[0].ToString();
                //product.Name = data[1].ToString();
                // list.Add(product);
                Console.Write(data[0] + " ");
                Console.Write(data[1] + " ");
                Console.Write(data[2] + " ");
                Console.Write(data[3] + " ");

            }

            //foreach (var pr in list)
            //{
            //  Console.WriteLine(pr.Id +"-"+pr.Name);
            //}
            conn.Close();
        }
    }

}

    

