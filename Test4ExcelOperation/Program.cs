using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Test4ExcelOperation
{
    class Program
    {
        static void Main(string[] args)
        {
            while(true)
            {
                try
                {
                    Console.WriteLine("请输入Excel文件路径或直接拖入Excel文件(输入exit则退出程序)：");
                    string filename = Console.ReadLine();
                    if (filename.ToUpper() == "EXIT")
                    {
                        break;
                    }
                    ExcelOperation operation = new ExcelOperation();
                    operation.Excel2Docx(filename);
                    Console.WriteLine("表格转换完成！");
                }
                catch(Exception err)
                {
                    Console.WriteLine(err.ToString());
                }
            }
            
        }
    }
}
