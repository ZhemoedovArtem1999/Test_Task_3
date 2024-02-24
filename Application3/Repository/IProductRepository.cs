using Application3.Models;
using Application3.Repository.Filter;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Application3.Repository
{
    public interface IProductRepository
    {
        IEnumerable<Product> Select(ProductFilter filter);
    }
}
