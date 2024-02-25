using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Application3.Service
{
    public class ValidateModel
    {
        public bool IsValid {get; set;}
        public List<string> Errors { get; set;}

        public ValidateModel()
        {
            Errors = new();
        }
    }
}
