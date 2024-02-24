using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Application3.Models
{
    public class Order
    {
        public int Id { get; set; }
        public int ProductId { get; set; }
        public int ClientId { get; set; }
        public int Number {  get; set; }
        public int RequiredQuantity { get; set; }
        public DateTime DatePlacement { get; set; }
    }
}
