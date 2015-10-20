using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Models;

namespace Data
{
    public class BooksDbContext: DbContext
    {
        public BooksDbContext():base("BooksDB")
        {
        }

        public virtual IDbSet<Book> Books { get; set; }
        public virtual IDbSet<BookPublisher> BookPublishers { get; set; }
        public virtual IDbSet<Author> Authors { get; set; }
        public virtual IDbSet<Sale> Sales { get; set; }
    }
}
