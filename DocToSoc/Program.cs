using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DocToSoc
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            try
            {
                new DocToSoc.DocSoc();
            }
            catch (Exception e)
            {
                Console.WriteLine("DocToSoc Failed");
                Console.WriteLine(e.ToString());
            }
        }
    }
}
