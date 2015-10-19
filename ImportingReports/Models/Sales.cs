using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Models
{
    public class Sales
    {
        public int SalesId { get; set; }

        [Required]
        [DataType(DataType.DateTime)]
        public DateTime date { get; set; }

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
    }
}
