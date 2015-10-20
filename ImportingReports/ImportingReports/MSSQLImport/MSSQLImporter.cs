namespace ImportingReports.MSSQLImport
{
    using System;
    using System.Collections;
    using System.IO;

    using Models;
    using Newtonsoft.Json;

    public class MSSQLImporter
    {
        public static void GetBookInfo(IList sales, string path)
        {
            using (StreamReader file = File.OpenText(path))
            {
                JsonSerializer serializer = new JsonSerializer();
                Sale sale = (Sale)serializer.Deserialize(file, typeof(Sale));

                sales.Add(sale);

                Console.WriteLine("Read!");
            }
        }
    }
}
