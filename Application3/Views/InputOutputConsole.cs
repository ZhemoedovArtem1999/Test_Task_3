using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Application3.Views
{
    public class InputOutputConsole : IInput, IOutput
    {
        public string Input(string message)
        {
            Console.Write(message + " - ");
            return Console.ReadLine();
        }

        public void Output(string message)
        {
            Console.WriteLine(message);
        }
    }
}
