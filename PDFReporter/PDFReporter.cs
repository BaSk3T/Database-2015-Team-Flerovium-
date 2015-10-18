namespace PDFReporter
{
    using System.Collections.Generic;
    using System.Data.Entity;
    using System.IO;
    using iTextSharp.text;
    using iTextSharp.text.pdf;

    public class PDFReporter
    {
        private const string Path = @"../../../Output/";

        public PDFReporter(DbContext db)
        {
            this.DataBase = db;
        }

        public DbContext DataBase { get; set; }

        public void GeneratePdfEmployeesReport(string documentName, List<Employees> listOfEmployees)
        {
            var headerFontStyle = PDFStyle.SetFontStyle("Arial", 14, Font.BOLD);
            var employeesHeadersFontStyle = PDFStyle.SetFontStyle("Arial", 10, Font.BOLD);
            var cellsFontStyle = PDFStyle.SetFontStyle("Arial", 8, Font.NORMAL);

            var document = this.CreatePdfDocument(Path + documentName);
            document.Open();

            var table = this.CreatePdfTable(3, new int[] { 4, 4, 4 });

            var documentHeader = this.CreatePdfCell("Employees", headerFontStyle, PDFStyle.Color(160, 170, 180), 3, 1);

            var lastNameHeader = this.CreatePdfCell("Last Name", employeesHeadersFontStyle, PDFStyle.Color(200, 200, 200));
            var firstNameHeader = this.CreatePdfCell("First Name", employeesHeadersFontStyle, PDFStyle.Color(200, 200, 200));
            var titleHeader = this.CreatePdfCell("Title", employeesHeadersFontStyle, PDFStyle.Color(200, 200, 200));
            
            table.AddCell(documentHeader);
            table.AddCell(lastNameHeader);
            table.AddCell(firstNameHeader);
            table.AddCell(titleHeader);

            foreach (var employee in listOfEmployees)
            {
                var lastName = this.CreatePdfCell(employee.LastName, cellsFontStyle, PDFStyle.Color(255, 255, 255));
                var firstName = this.CreatePdfCell(employee.FirstName, cellsFontStyle, PDFStyle.Color(255, 255, 255));
                var title = this.CreatePdfCell(employee.Title, cellsFontStyle, PDFStyle.Color(255, 255, 255));
            
                table.AddCell(lastName);
                table.AddCell(firstName);
                table.AddCell(title);
            }

            document.Add(table);
            document.Close();
        }

        private Document CreatePdfDocument(string documentName)
        {
            // Create a Document object
            var document = new Document(PageSize.A4, 50, 50, 25, 25);

            // Create a new PdfWriter object, specifying the output stream
            var output = new FileStream(documentName + ".pdf", FileMode.Create);
            var writer = PdfWriter.GetInstance(document, output);

            return document;
        }

        private PdfPTable CreatePdfTable(int tableCols, int[] parameters)
        {
            var infoTable = new PdfPTable(tableCols);

            infoTable.HorizontalAlignment = 1;
            infoTable.SpacingBefore = 10;
            infoTable.SpacingAfter = 10;
            infoTable.DefaultCell.Border = 1;
            infoTable.SetWidths(parameters);

            return infoTable;
        }

        private PdfPCell CreatePdfCell(string cellInfo, Font font, BaseColor backgroundColor, int colSpan = 0, int horizontalAligment = 0)
        {
            var cell = new PdfPCell(new Phrase(cellInfo, font));
            cell.BackgroundColor = backgroundColor;
            cell.Colspan = colSpan;
            cell.HorizontalAlignment = horizontalAligment;
            return cell;
        }       
    }
}
