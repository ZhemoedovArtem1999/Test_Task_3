using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Application3.Views
{
    public class OutputException : IOutput
    {
        public void Output(string message)
        {
            
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine("Исключение: " + message);
            Console.ResetColor();
        }
    }
}
