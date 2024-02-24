using Application3.HelpersExcel;
using Application3.Models;
using Application3.Repository.Filter;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Application3.Repository.ExcelRepository
{
    public class ExcelOrderRepository : IOrderRepository
    {
        private string _filePath;
        private string _sheetName = "Заявки";

        private List<string> _titleHeader = new List<string>()
        {
            "Код заявки", "Код товара", "Код клиента", "Номер заявки", "Требуемое количество", "Дата размещения"
        };
        private Dictionary<string, int> _indexHeader = new();

        public ExcelOrderRepository(string filePath)
        {
            _filePath = filePath;

        }

        private List<Order> ApplyFilter(IEnumerable<Order> entities, OrderFilter filter)
        {
            if (filter.ProductId != null) entities = entities.Where(order => order.ProductId == filter.ProductId);


            if (filter.Year != null && filter.Month != null) entities = entities.Where(order => order.DatePlacement.Year == filter.Year && order.DatePlacement.Month == filter.Month);

            return entities.ToList();
        }


        public IEnumerable<Order> Select(OrderFilter filter)
        {
            ExcelData excelDataOrder = new HelperExcel().ReadExcel(_filePath, _sheetName);
            if (_indexHeader.Count == 0)
                foreach (string title in _titleHeader)
                {
                    _indexHeader.Add(title, excelDataOrder.Headers.IndexOf(title));
                }
            List<Order> orders = new();

            foreach (var row in excelDataOrder.DataRows)
            {
                if (row[_indexHeader.GetValueOrDefault(_titleHeader[0])] != "")
                    orders.Add(new Order
                    {


                        Id = int.Parse(row[_indexHeader.GetValueOrDefault(_titleHeader[0])]),
                        ProductId = int.Parse(row[_indexHeader.GetValueOrDefault(_titleHeader[1])]),
                        ClientId = int.Parse(row[_indexHeader.GetValueOrDefault(_titleHeader[2])]),
                        Number = int.Parse(row[_indexHeader.GetValueOrDefault(_titleHeader[3])]),
                        RequiredQuantity = int.Parse(row[_indexHeader.GetValueOrDefault(_titleHeader[4])]),
                        DatePlacement = DateTime.Parse(row[_indexHeader.GetValueOrDefault(_titleHeader[5])]),


                    });
            }


            //Product product = excelDataProduct.DataRows.Select(row => new Product
            //{
            //    Id = int.Parse(row[_indexHeader.GetValueOrDefault(_titleHeader[0])]),
            //    Title = (row[_indexHeader.GetValueOrDefault(_titleHeader[1])]),
            //    Unit = (row[_indexHeader.GetValueOrDefault(_titleHeader[2])]),
            //    Price = double.Parse(row[_indexHeader.GetValueOrDefault(_titleHeader[3])])

            //}).Where(product => product.Title == filter.Title).FirstOrDefault();


            return this.ApplyFilter(orders, filter);



        }


    }
}
