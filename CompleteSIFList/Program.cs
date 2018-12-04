using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Spire.Xls;

namespace CompleteSIFList
{
    class Program
    {
        static void Main(string[] args)
        {
            while(true)
            {
                Console.WriteLine("请输入文件路径或将文件拖入窗口：");
                string filename = Console.ReadLine();
                if(filename.ToUpper() == "EXIT")
                {
                    return;
                }
                Operation(filename.Replace("\"", ""));
                Console.WriteLine("转换完毕！");
            }
        }

        static void Operation(string filename)
        {
            Workbook excel = new Workbook();
            excel.LoadFromFile(filename);
            int i = 1;
            foreach(Worksheet sheet in excel.Worksheets )
            {
                if (sheet.Name == "清单" || sheet.Name == "序列引用")
                {
                    i++;
                    continue;
                }

                excel.Worksheets["清单"].Range["A" + i.ToString()].Value = (i-1).ToString();
                excel.Worksheets["清单"].Range["B" + i.ToString()].Value = "SIF" + sheet.Range["A5"].Value;
                excel.Worksheets["清单"].Range["C" + i.ToString()].Value = sheet.Range["B5"].Value;

                excel.Worksheets["清单"].Range["E" + i.ToString()].Value = sheet.Name;
                i++;
            }

            excel.SaveToFile(filename.Replace(".xlsx", "_result.xlsx").Replace(".xls", "_result.xlsx").Replace("_result.xlsxx", ".xlsx"), FileFormat.Version2013);
            
        }
    }
}
