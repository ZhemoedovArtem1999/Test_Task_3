using Application3.ModelView;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Application3.Models
{
    public class ClientProduct
    {
        public Status Status { get; set; }
        public Client Client { get; set; }
        public int CountProduct { get; set; }
        public double? PriceProduct { get; set; }
        public DateTime Date { get; set; }

        public override string ToString()
        {
            return $"Id клиента - {Client.Id} " +
                $"Наименование организации - {Client.OrganizationTitle} " +
                $"Контактное лицо - {Client.ContactPerson} " +
                $"Количество - {CountProduct} " +
                $"Цена - {PriceProduct} " +
                $"Дата - {Date.ToString("dd.MM.yyyy")} "
                ;
        }
    }
}
