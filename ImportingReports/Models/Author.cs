using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Models
{
    public class Author
    {
        private ICollection<Book> books;

        public Author()
        {
            this.books = new HashSet<Book>();
        }

        public int AuthorId { get; set; }

        [Required]
        [DataType(DataType.Text)]
        public string AuthorName { get; set; }

        public virtual ICollection<Book> Books
        {
            get { return this.books; }
            set { this.books = value; }
        }
    }
}
