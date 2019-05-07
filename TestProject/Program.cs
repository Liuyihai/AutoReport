using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Spire.Xls;

namespace TestProject
{
    class Program
    {
        static void Main(string[] args)
        {
            while(true)
            {
                Console.WriteLine("请拖入文件：");
                string path = Console.ReadLine();
                Class2 class2 = new Class2();
                class2.DataDeal(path);
            }
        }
        
    }
}
