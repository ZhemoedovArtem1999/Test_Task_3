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
    public class ExcelClientRepository : IClientRepository
    {
        private string _filePath;
        private string _sheetName = "Клиенты";

        private List<string> _titleHeader = new List<string>()
        {
            "Код клиента", "Наименование организации", "Адрес", "Контактное лицо (ФИО)"
        };
        private Dictionary<string, int> _indexHeader = new();

        public ExcelClientRepository(string filePath)
        {
            _filePath = filePath;

        }

        private List<Client> ApplyFilter(IEnumerable<Client> entities, ClientFilter filter)
        {
            if (filter.Id != null)
            {
                entities = entities.Where(client => client.Id == filter.Id);
            }

            return entities.ToList();
        }


        public IEnumerable<Client> Select(ClientFilter filter)
        {
            ExcelData excelDataClient = new HelperExcel().ReadExcel(_filePath, _sheetName);
            if (_indexHeader.Count == 0)
                foreach (string title in _titleHeader)
                {
                    _indexHeader.Add(title, excelDataClient.Headers.IndexOf(title));
                }
            List<Client> clients = new();

            foreach (var row in excelDataClient.DataRows)
            {
                if (row[_indexHeader.GetValueOrDefault(_titleHeader[0])] != "")
                    clients.Add(new Client
                    {


                        Id = int.Parse(row[_indexHeader.GetValueOrDefault(_titleHeader[0])]),
                        OrganizationTitle = (row[_indexHeader.GetValueOrDefault(_titleHeader[1])]),
                        Address = (row[_indexHeader.GetValueOrDefault(_titleHeader[2])]),
                        ContactPerson = (row[_indexHeader.GetValueOrDefault(_titleHeader[3])]),

                    });
            }


            //Product product = excelDataProduct.DataRows.Select(row => new Product
            //{
            //    Id = int.Parse(row[_indexHeader.GetValueOrDefault(_titleHeader[0])]),
            //    Title = (row[_indexHeader.GetValueOrDefault(_titleHeader[1])]),
            //    Unit = (row[_indexHeader.GetValueOrDefault(_titleHeader[2])]),
            //    Price = double.Parse(row[_indexHeader.GetValueOrDefault(_titleHeader[3])])

            //}).Where(product => product.Title == filter.Title).FirstOrDefault();

            return this.ApplyFilter(clients, filter);




        }

        public Client GetById(int id)
        {
            return this.Select(new ClientFilter() { Id = id }).FirstOrDefault();
        }


        public void UpdateClient(string organizationTitle, string newContactPerson)
        {
            HelperExcel helperExcel = new HelperExcel();
            List<string> header = helperExcel.GetHeader(_filePath, _sheetName);
            if (_indexHeader.Count == 0)
                foreach (string title in _titleHeader)
                {
                    _indexHeader.Add(title, header.IndexOf(title));
                }

            int columnIndexOrganizationTitle = _indexHeader.GetValueOrDefault(_titleHeader[1]);
            int columnIndexContactPerson = _indexHeader.GetValueOrDefault(_titleHeader[3]);

            helperExcel.UpdateExcel(_filePath, _sheetName, columnIndexOrganizationTitle, columnIndexContactPerson, organizationTitle, newContactPerson);




        }
    }
}
