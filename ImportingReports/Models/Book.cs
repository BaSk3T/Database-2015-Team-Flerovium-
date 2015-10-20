using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Security.Policy;
using System.Text;
using System.Threading.Tasks;

namespace Models
{
    public class Book
    {
        private ICollection<Sale> sales;

        public Book()
        {
            this.sales = new HashSet<Sale>();
        }
        
        public int BookId { get; set; }

        [Required]
        [DataType(DataType.Text)]
        public string Title { get; set; }

        [Required]
        public int AuthorId { get; set; }

        [Required]
        public int BookPublisherId { get; set; }

        [Required]
        [DataType(DataType.Currency)]
        public decimal BasePrice { get; set; }

        public virtual Author Author { get; set; }

        public virtual BookPublisher Publisher { get; set; }

        public virtual ICollection<Sale> Sales {
            get { return this.sales; }
            set { this.sales = value; }
        }
    }
}
