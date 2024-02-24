using Application3.HelpersExcel;
using Application3.Models;
using Application3.Repository.Filter;
using DocumentFormat.OpenXml.Office2016.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection.Metadata.Ecma335;
using System.Text;
using System.Threading.Tasks;

namespace Application3.Repository.ExcelRepository
{
    public class ExcelProductRepository : IProductRepository
    {
        private string _filePath;
        private string _sheetName = "Товары";

        private List<string> _titleHeader = new List<string>()
        {
            "Код товара", "Наименование", "Ед. измерения", "Цена товара за единицу"
        };
        private Dictionary<string, int> _indexHeader = new();

        public ExcelProductRepository(string filePath)
        {
            _filePath = filePath;

        }

        private List<Product> ApplyFilter(IEnumerable<Product> entities, ProductFilter filter)
        {
            if (filter.Title != null)
            {
                entities = entities.Where(order => order.Title == filter.Title);
            }

            return entities.ToList();
        }



        public IEnumerable<Product> Select(ProductFilter filter)
        {
            ExcelData excelDataProduct = new HelperExcel().ReadExcel(_filePath, _sheetName);
            if (_indexHeader.Count == 0)
                foreach (string title in _titleHeader)
                {
                    _indexHeader.Add(title, excelDataProduct.Headers.IndexOf(title));
                }
            List<Product> products = new();

            foreach (var row in excelDataProduct.DataRows)
            {
                products.Add(new Product
                {
                    Id = int.Parse(row[_indexHeader.GetValueOrDefault(_titleHeader[0])]),
                    Title = (row[_indexHeader.GetValueOrDefault(_titleHeader[1])]),
                    Unit = (row[_indexHeader.GetValueOrDefault(_titleHeader[2])]),
                    Price = double.Parse(row[_indexHeader.GetValueOrDefault(_titleHeader[3])])

                });
            }


            //Product product = excelDataProduct.DataRows.Select(row => new Product
            //{
            //    Id = int.Parse(row[_indexHeader.GetValueOrDefault(_titleHeader[0])]),
            //    Title = (row[_indexHeader.GetValueOrDefault(_titleHeader[1])]),
            //    Unit = (row[_indexHeader.GetValueOrDefault(_titleHeader[2])]),
            //    Price = double.Parse(row[_indexHeader.GetValueOrDefault(_titleHeader[3])])

            //}).Where(product => product.Title == filter.Title).FirstOrDefault();


            return this.ApplyFilter(products, filter);



        }



    }
}
