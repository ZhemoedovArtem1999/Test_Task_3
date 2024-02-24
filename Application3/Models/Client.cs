using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Application3.Models
{
    public class Client
    {
        public int Id { get; set; }
        public string OrganizationTitle { get; set; }
        public string Address { get; set; }
        public string ContactPerson { get; set; }

        public override string ToString()
        {
            return $"Наименование организации - {OrganizationTitle} " +
                $"Адрес - {Address} " +
                $"Контактное лицо - {ContactPerson} ";
        }
    }
}
