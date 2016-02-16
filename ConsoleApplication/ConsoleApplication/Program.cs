using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("hello");
            var currentDay = DateTime.Now.AddDays(-60);
            Console.WriteLine(currentDay);
        }
    }
}
