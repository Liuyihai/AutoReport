using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
//using Spire.Doc;
//using Spire.Doc.Documents;
using Spire.Xls;

namespace OutInfo
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
                    GetData(filename.Replace("\"", ""));
                    Console.WriteLine("表格转换完成！");
                }
                catch (Exception err)
                {
                    Console.WriteLine(err.ToString());
                }
            }
        }

        static void GetData(string filename)
        {
            Workbook wb = new Workbook();
            wb.LoadFromFile(filename);
            Workbook result = new Workbook();
            Worksheet resws = result.CreateEmptySheet("result");
            resws.Range["A1"].Value = "节点编号";
            resws.Range["B1"].Value = "编号";
            resws.Range["C1"].Value = "SIF回路";
            resws.Range["D1"].Value = "后果描述";
            resws.Range["E1"].Value = "SIF描述";
            resws.Range["F1"].Value = "SIL等级要求";
            resws.Range["G1"].Value = "PID";
            resws.Range["H1"].Value = "保护层";

            int count = 2;
            foreach(Worksheet ws in wb.Worksheets)
            {
                for(int i = 1;i <= ws.Rows.Count();i++)
                {
                    if(ws.Range["K"+i.ToString()].Value == "编号")
                    {
                        resws.Range["A" + count.ToString()].Value = ws.Range["H" + i.ToString()].Value;//节点编号
                        resws.Range["B" + count.ToString()].Value = ws.Range["L" + i.ToString()].Value;//编号
                        resws.Range["C" + count.ToString()].Value = ws.Range["AB" + i.ToString()].Value;//SIF回路
                        resws.Range["D" + count.ToString()].Value = ws.Range["L" + (i + 1).ToString()].Value;//后果描述
                        resws.Range["E" + count.ToString()].Value = ws.Range["AB" + (i + 1).ToString()].Value;//SIF描述
                        resws.Range["F" + count.ToString()].Value = ws.Range["B" + (i + 3).ToString()].Value;//SIL等级要求
                        resws.Range["G" + count.ToString()].Value = ws.Range["H" + (i + 2).ToString()].Value;//PID
                        resws.Range["H" + count.ToString()].Value = ws.Range["I" + (i + 6).ToString()].Value;//保护层
                        Console.WriteLine(resws.Range["A" + count.ToString()].Value + "\t" +
                            resws.Range["B" + count.ToString()].Value + "\t" +
                            resws.Range["C" + count.ToString()].Value + "\t" +
                            resws.Range["D" + count.ToString()].Value + "\t" +
                            resws.Range["E" + count.ToString()].Value + "\t" +
                            resws.Range["F" + count.ToString()].Value + "\t" +
                            resws.Range["G" + count.ToString()].Value + "\t" +
                            resws.Range["H" + count.ToString()].Value );
                        count++;
                    }
                }
            }

            result.SaveToFile(filename.Replace(".xlsx", "_result.xlsx").Replace(".xls", "_result.xlsx").Replace("_result.xlsxx", ".xlsx"),ExcelVersion.Version2013);
        }
    }
}
