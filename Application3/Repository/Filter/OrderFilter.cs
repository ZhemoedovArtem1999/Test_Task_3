using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Application3.Repository.Filter
{
    public class OrderFilter
    {
        public int? ProductId {  get; set; } 
        public int? Year { get; set; }
        public int? Month { get; set;}
    }
}
