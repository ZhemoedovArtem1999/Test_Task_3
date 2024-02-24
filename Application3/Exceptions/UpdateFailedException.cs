using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Application3.Exceptions
{
    public class UpdateFailedException:Exception
    {
        public UpdateFailedException(string message) : base(message) { }
    }
}
