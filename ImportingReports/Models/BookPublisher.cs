using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace Models
{
    public class BookPublisher
    {
        private ICollection<Book> books;
        public BookPublisher()
        {
            this.books = new HashSet<Book>();
        }
        
        public int BookPublisherId { get; set; }

        [Required]
        [DataType(DataType.Text)]
        public string BookPublisherName { get; set; }

        public virtual ICollection<Book> Books { get { return this.books; } set { this.books = value; } }
    }
}
