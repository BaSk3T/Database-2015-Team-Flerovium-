using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using MongoDB.Bson;

namespace ImportingReports
{
    public class BookMongo
    {

        public static Expression<Func<XElement, BookMongo>> FromXElement
        {
            get
            {
                return el => new BookMongo
                {
                    Id = el.Attribute("id").ToString(),
                    AuthorId = el.Attribute("author-id").ToString(),
                    PublisherId = el.Attribute("publisher-id").ToString(),
                    Title = el.Element("title").Value,
                    Author = el.Element("author").Value,
                    Publisher = el.Element("publisher").Value,
                    BasePrice = el.Element("basePrice").Value
                };
            }
        }

        public string BasePrice { get; set; }

        public string Publisher { get; set; }

        public string Author { get; set; }

        public string Title { get; set; }

        public string PublisherId { get; set; }

        public string AuthorId { get; set; }

        public string Id { get; set; }

        public override string ToString()
        {
            return string.Format("Id: {0}, AuthorId: {1}, Title: {2}", this.Id, this.AuthorId, this.Title);
        }

        public static Expression<Func<BookMongo, BsonDocument>> ToBsonDocument
        {
            get { return book => BookMongo.ParseToBsonDocument(book); }
        }

        private static BsonDocument ParseToBsonDocument(BookMongo book)
        {
            var doc = new BsonDocument();
            doc["Title"] = book.Title;
            doc["BookId"] = book.Id;
            doc["Author"] = book.Author;
            doc["AuthorId"] = book.AuthorId;
            doc["Publisher"] = book.Publisher;
            doc["PublisherId"] = book.PublisherId;
            doc["BasePrice"] = book.BasePrice;

            return doc;
        }
    }
}
