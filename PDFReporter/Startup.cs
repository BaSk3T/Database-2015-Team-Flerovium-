namespace PDFReporter
{
    using System;
    using System.Linq;

    public class Startup
    {
        public static void Main()
        {
            var db = new NorthwindEntities();
            
            var testPdf = new PDFReporter(db);

            var data = db.Employees.ToList();

            testPdf.GeneratePdfEmployeesReport("Employees", data);

            Console.WriteLine("Report created.");
        }
    }
}
