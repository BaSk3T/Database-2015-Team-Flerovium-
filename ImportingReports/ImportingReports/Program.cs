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
using System.Data.Entity.Migrations;
using System.Security.Permissions;
using Data;
using Models;
using Models.MySQLModels;
using MySql.Data;
using MySql.Data.MySqlClient;
using OfficeOpenXml;


namespace ImportingReports
{
    class Program
    {
        private static void Main()
        {
            Run();
            Console.ReadKey();
        }

        static async void Run()
        {

            //var extractPath = ExtractingFromZip();
            var xmlPath =
                @"..\..\XmlImportFile\Books.xml";
            XDocument doc = XDocument.Load(xmlPath);

            var zipFilePath = @"C:\Users\Cookie\Desktop\Telerik Academy\Databases\Database-2015-Team-Flerovium-\ImportingReports\ImportingReports\ZipFile\Sales-Reports.zip";
            var extractedFolderPath = ExtractZip(zipFilePath);
            var dateNamedFolders = GetDateNamedFolders(extractedFolderPath);

            var allSales = new List<Sale>();

            foreach (var folder in dateNamedFolders)
            {
                foreach (var file in folder.GetFiles())
                {
                    var salesForFile = ExtractExcelData(file.FullName, file.Name, folder.Name);
                    allSales.AddRange(salesForFile);
                }
            }

            allSales.ForEach(Console.WriteLine);

            var root = doc.Root;

            var books = root.Elements().AsQueryable().Select(BookMongo.FromXElement).ToList();

            var client = new MongoClient("mongodb://localhost");
            var db = client.GetDatabase("books-db");

            var booksCollection = db.GetCollection<BsonDocument>("books");

            var bsonBooks = books.AsQueryable().Select(BookMongo.ToBsonDocument);

            await booksCollection.InsertManyAsync(bsonBooks);

            var dbContext = new BooksDbContext();
            var mongoBooks = await booksCollection.Find(new BsonDocument()).ToListAsync();
            var bookModels = mongoBooks.Select(bookBson =>
            {
                var authorIdToString = bookBson["AuthorId"].ToString();
                var author =
                    dbContext.Authors.FirstOrDefault(a => a.AuthorId.ToString() == authorIdToString);
                if (author == null)
                {
                    author = new Author
                    {
                        AuthorId = int.Parse(bookBson["AuthorId"].ToString()),
                        AuthorName = bookBson["Author"].ToString()
                    };
                    dbContext.Authors.AddOrUpdate(a => a.AuthorName, author);
                    dbContext.SaveChanges();
                }

                var publisherIdToString = bookBson["PublisherId"].ToString();
                var publisher = dbContext.BookPublishers.FirstOrDefault(
                     b => b.BookPublisherId.ToString() == publisherIdToString);

                if (publisher == null)
                {
                    publisher = new BookPublisher
                    {
                        BookPublisherId = int.Parse(bookBson["PublisherId"].ToString()),
                        BookPublisherName = bookBson["Publisher"].ToString()
                    };
                    dbContext.BookPublishers.AddOrUpdate(p => p.BookPublisherName, publisher);
                    dbContext.SaveChanges();
                }

                var book = new Book
                {
                    BookId = int.Parse(bookBson["BookId"].ToString()),
                    AuthorId = int.Parse(bookBson["AuthorId"].ToString()),
                    BookPublisherId = int.Parse(bookBson["PublisherId"].ToString()),
                    Title = bookBson["Title"].ToString(),
                    BasePrice = decimal.Parse(bookBson["BasePrice"].ToString()),
                    Author = author,
                    Publisher = publisher
                };
                return book;
            });

            Console.WriteLine("--------------- {0} Saving Books", bookModels.Count());
            dbContext.Books.AddOrUpdate(b => b.Title, bookModels.ToArray());

            dbContext.SaveChanges();
            Console.WriteLine("---------------Books saved!");

            allSales.ForEach(sale => dbContext.Sales.Add(sale));
            dbContext.SaveChanges();

            Console.ReadKey();
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




        private static IEnumerable<Sale> ExtractExcelData(string excelFilePath, string excelFileName, string date)
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
            var sales = new List<Sale>();

            foreach (DataRow data in table.Rows)
            {
                try
                {
                    var sale = new Sale
                    {
                        BookId = int.Parse(data[0].ToString()),
                        Quantity = int.Parse(data[1].ToString()),
                        UnitPrice = decimal.Parse(data[2].ToString()),
                        Sum = decimal.Parse(data[3].ToString()),
                        Location = excelFileName,
                        Date = DateTime.Parse(date)
                    };

                    sales.Add(sale);
                }
                catch
                {
                }
            }

            conn.Close();
            return sales;
        }

        private static ICollection<Courier> MySQLReadCouriersTable()
        {
            var server = "localhost";
            var database = "specialdeliveries";
            var uid = "root";
            var password = "1234";
            string connectionString = "SERVER=" + server + ";" + "DATABASE=" +
            database + ";" + "UID=" + uid + ";" + "PASSWORD=" + password + ";";

            MySqlConnection connection = new MySqlConnection(connectionString);
            connection.Open();
            var query = "SELECT * FROM Couriers";

            var command = new MySqlCommand(query, connection);

            MySqlDataReader reader = command.ExecuteReader();
            ICollection<Courier> courierCollection = new List<Courier>();


            while (reader.Read())
            {
                var currentCourier = new Courier()
                {
                    Id = (int)reader.GetValue(0),
                    Name = (string)reader.GetValue(1),
                    TownId = (int)reader.GetValue(2)
                };
                courierCollection.Add(currentCourier);
            }

            return courierCollection;
        }

        private static ICollection<Town> MySQLReadTownsTable()
        {
            var server = "localhost";
            var database = "specialdeliveries";
            var uid = "root";
            var password = "1234";
            string connectionString = "SERVER=" + server + ";" + "DATABASE=" +
            database + ";" + "UID=" + uid + ";" + "PASSWORD=" + password + ";";

            MySqlConnection connection = new MySqlConnection(connectionString);
            connection.Open();
            var query = "SELECT * FROM Towns";

            var command = new MySqlCommand(query, connection);

            MySqlDataReader reader = command.ExecuteReader();
            ICollection<Town> townsCollection = new List<Town>();


            while (reader.Read())
            {
                var currentCourier = new Town()
                {
                    Id = (int)reader.GetValue(0),
                    Name = (string)reader.GetValue(1),
                };
                townsCollection.Add(currentCourier);
            }

            return townsCollection;
        }

        private static ICollection<Product> MySQLReadProductsTable()
        {
            var server = "localhost";
            var database = "specialdeliveries";
            var uid = "root";
            var password = "1234";
            string connectionString = "SERVER=" + server + ";" + "DATABASE=" +
            database + ";" + "UID=" + uid + ";" + "PASSWORD=" + password + ";";

            MySqlConnection connection = new MySqlConnection(connectionString);

            connection.Open();

            var query = "SELECT * FROM Products";

            var command = new MySqlCommand(query, connection);

            MySqlDataReader reader = command.ExecuteReader();
            var productsCollection = new List<Product>();


            while (reader.Read())
            {
                var currentCourier = new Product()
                {
                    Id = (int)reader.GetValue(0),
                    Name = (string)reader.GetValue(1),
                    Description = (string)reader.GetValue(2),
                    CourierId = (int)reader.GetValue(3)
                };
                productsCollection.Add(currentCourier);
            }
            
            return productsCollection;
        }

        public static void WriteExcelTownSheet(ICollection<Town> towns, string fileName)
        {
            // string fileName = "Towns.xlsx";
            var file = new FileInfo(fileName);
            var rowNumber = 2;

            using (var package = new ExcelPackage(file))
            {

                ExcelWorksheet TownSheet = package.Workbook.Worksheets.Add("Townss");
                TownSheet.Cells[1, 1].Value = "Id";
                TownSheet.Cells[1, 2].Value = "Name";

                using (var range = TownSheet.Cells[1, 1, 1, 2])
                {
                    range.Style.Font.Bold = true;
                    range.Style.ShrinkToFit = false;
                }


                foreach (var town in towns)
                {
                    TownSheet.Cells[rowNumber, 1].Value = town.Id;
                    TownSheet.Cells[rowNumber, 2].Value = town.Name;
                    rowNumber++;
                }
                package.Save();
            }
        }

        public static void WriteExcelCourierSheet(ICollection<Courier> couriers, string fileName)
        {
            var file = new FileInfo(fileName);
            var rowNumber = 2;

            using (var package = new ExcelPackage(file))
            {

                ExcelWorksheet CourierSheet = package.Workbook.Worksheets.Add("Couriers");
                CourierSheet.Cells[1, 1].Value = "Id";
                CourierSheet.Cells[1, 2].Value = "Name";
                CourierSheet.Cells[1, 3].Value = "TownId";


                using (var range = CourierSheet.Cells[1, 1, 1, 3])
                {
                    range.Style.Font.Bold = true;
                    range.Style.ShrinkToFit = false;
                }


                foreach (var courier in couriers)
                {
                    CourierSheet.Cells[rowNumber, 1].Value = courier.Id;
                    CourierSheet.Cells[rowNumber, 2].Value = courier.Name;
                    CourierSheet.Cells[rowNumber, 3].Value = courier.TownId;
                    rowNumber++;
                }
                package.Save();
            }
        }

        public static void WriteExcelProductSheet(ICollection<Product> products, string fileName)
        {
            var file = new FileInfo(fileName);
            var rowNumber = 2;

            using (var package = new ExcelPackage(file))
            {

                ExcelWorksheet ProductSheet = package.Workbook.Worksheets.Add("Products");
                ProductSheet.Cells[1, 1].Value = "Id";
                ProductSheet.Cells[1, 2].Value = "Name";
                ProductSheet.Cells[1, 3].Value = "Description";
                ProductSheet.Cells[1, 4].Value = "CourierId";

                using (var range = ProductSheet.Cells[1, 2, 1, 4])
                {
                    range.Style.Font.Bold = true;
                    range.Style.ShrinkToFit = false;
                }


                foreach (var product in products)
                {
                    ProductSheet.Cells[rowNumber, 1].Value = product.Id;
                    ProductSheet.Cells[rowNumber, 2].Value = product.Name;
                    ProductSheet.Cells[rowNumber, 3].Value = product.Description;
                    ProductSheet.Cells[rowNumber, 4].Value = product.CourierId;
                    rowNumber++;
                }

                package.Save();
            }
        }
    }
}



