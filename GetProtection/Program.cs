using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Spire.Xls;
using System.IO;

namespace GetProtection
{
    class Program
    {
        static void Main(string[] args)
        {
            while (true)
            {
                try
                {
                    Console.WriteLine("请输入Excel文件路径或直接拖入Excel文件(输入exit则退出程序)：");
                    string filename = Console.ReadLine();
                    if (filename.ToUpper() == "EXIT")
                    {
                        break;
                    }
                    filename = filename.Replace("\"", "");
                    List<string> protection = Protection(filename);
                    FileStream file = new FileStream(filename.Replace(".xlsm", ".txt").Replace(".xlsx", ".txt").Replace(".xls", ".txt"), FileMode.Create);
                    file.Seek(0, SeekOrigin.End);
                    foreach (string p in protection)
                    {
                        Console.WriteLine(p);
                        byte[] fw = System.Text.Encoding.Default.GetBytes(p + "\r\n");
                        file.Write(fw, 0, fw.Length);
                        file.Seek(0, SeekOrigin.End);
                    }
                    file.Close();
                    Console.WriteLine("保护层提取完成！");
                }
                catch (Exception err)
                {
                    Console.WriteLine(err.ToString());
                }
            }
        }

        private static List<string> Protection(string filepath)
        {
            Workbook excel = new Workbook();
            excel.LoadFromFile(filepath);
            string protection = string.Empty;
            List<string> p = new List<string>();
            foreach (Worksheet sheet in excel.Worksheets)
            {
                Regex regex = new Regex(@"SIF List");
                if (sheet == excel.Worksheets[0] || sheet.Name == "SIL decision matrix")
                    continue;
                else if (regex.Match(sheet.Name).Success) continue;
                else
                {
                    if (sheet.Range["B46"].FormulaValue != null)
                        if (!p.Contains(sheet.Range["B46"].FormulaValue.ToString()))
                            p.Add(sheet.Range["B46"].FormulaValue.ToString());
                }
            }
            return p;
        }
    }
}
