using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Application3.Models;

namespace Application3.ModelView
{
    public class ClientProductView
    {
        public Status Status { get; set; }
        public List<ClientProduct> ClientProducts {  get; set; }
    }
}
