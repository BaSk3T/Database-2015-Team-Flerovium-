namespace ImportingReports
{
    using System.Collections.Generic;
    using System.Data.Entity;
    using System.IO;
    using iTextSharp.text;
    using iTextSharp.text.pdf;
    using Models;

    public class PDFReporter
    {
        private const string Path = @"";

        public PDFReporter(DbContext db)
        {
            this.DataBase = db;
        }

        public DbContext DataBase { get; set; }

        public void GeneratePdfSalesReport(string documentName, List<Sales> listOfSales)
        {
            var headerFontStyle = PDFStyle.SetFontStyle("Arial", 14, Font.BOLD);
            var salesHeadersFontStyle = PDFStyle.SetFontStyle("Arial", 10, Font.BOLD);
            var cellsFontStyle = PDFStyle.SetFontStyle("Arial", 8, Font.NORMAL);

            var document = this.CreatePdfDocument(Path + documentName);
            document.Open();

            var table = this.CreatePdfTable(3, new int[] { 4, 4, 4 });

            var documentHeader = this.CreatePdfCell("Sales", headerFontStyle, PDFStyle.Color(160, 170, 180), 3, 1);

            var bookIdHeader = this.CreatePdfCell("Book Id", salesHeadersFontStyle, PDFStyle.Color(200, 200, 200));
            var quantityHeader = this.CreatePdfCell("Quantity", salesHeadersFontStyle, PDFStyle.Color(200, 200, 200));
            var priceHeader = this.CreatePdfCell("Unit Price", salesHeadersFontStyle, PDFStyle.Color(200, 200, 200));

            table.AddCell(documentHeader);
            table.AddCell(bookIdHeader);
            table.AddCell(quantityHeader);
            table.AddCell(priceHeader);

            foreach (var sale in listOfSales)
            {
                var bookId = this.CreatePdfCell(sale.BookId.ToString(), cellsFontStyle, PDFStyle.Color(255, 255, 255));
                var quantity = this.CreatePdfCell(sale.Quantity.ToString(), cellsFontStyle, PDFStyle.Color(255, 255, 255));
                var price = this.CreatePdfCell(sale.UnitPrice.ToString(), cellsFontStyle, PDFStyle.Color(255, 255, 255));

                table.AddCell(bookId);
                table.AddCell(quantity);
                table.AddCell(price);
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
