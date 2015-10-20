using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Models
{
    public class Sale
    {
        public int SaleId { get; set; }

        [Required]
        [DataType(DataType.DateTime)]
        public DateTime Date { get; set; }

        [Required]
        public int BookId { get; set; }

        [Required]
        public int Quantity { get; set; }

        [Required]
        [DataType(DataType.Text)]
        public string Location { get; set; }

        [Required]
        public decimal UnitPrice { get; set; }

        [Required]
        public decimal Sum { get; set; }

        [Required]
        public decimal TotalSum { get; set; }

        public virtual Book Book { get; set; }

        public override string ToString()
        {
            return string.Format("{0}, {1}, {2}, {3}, {4}", this.SaleId, this.BookId, this.Quantity, this.Location, this.Date);
        }
    }
}
