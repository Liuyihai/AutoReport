using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Spire.Xls;

namespace Protection
{
    class Protection
    {
        static void Main(string[] args)
        {
            while(true)
            {
                Console.WriteLine("请输入文件路径：");
                string filepath = Console.ReadLine();
                if(ProFun(filepath.Replace("\"","")))
                {
                    Console.WriteLine("信息提取成功，请在原目录查看！");
                }
                else
                {
                    Console.WriteLine("啊哦~~~~~c出错啦！");
                }
            }
        }

        private static bool ProFun(string filename)
        {
            Workbook Lopa = new Workbook();
            Lopa.LoadFromFile(filename);
            Workbook result = new Workbook();
            Worksheet resultxls = result.Worksheets[0];
            try
            {
                int i = 1;
                foreach (Worksheet test in Lopa.Worksheets)
                {
                    Regex sheetname = new Regex(@"Sheet");
                    if (test.Name.ToUpper() == "SIF LIST" || test.Name == "结论" || test.Name == "清单" || test.Name.ToUpper() == "SIF LIST OLD" || test.Name == "序列引用" || sheetname.IsMatch(test.Name)) continue;

                    resultxls.Range["A" + i.ToString()].Value = test.Name;

                    foreach(var row in Lopa.Worksheets["清单"].Rows)
                    {
                        if(test.Name == row.Cells[1].Value)
                        {
                            resultxls.Range["B" + i.ToString()].Value = row.Cells[2].Value;
                        }
                    }



                    int p1 = int.Parse(test.Range["A1"].FormulaValue.ToString());
                    int p2 = int.Parse(test.Range["B1"].FormulaValue.ToString());
                    int p3 = int.Parse(test.Range["C1"].FormulaValue.ToString());
                    string[] col = new string[] { "I", "J", "K", "L", "M", };
                    List<string> protection = new List<string>();
                    for (int l = p1; l < p2 - 6; l += 2)
                    {
                        foreach (string c in col)
                        {
                            if (test.Range[c + l.ToString()].Value == null) continue;
                            Console.WriteLine(test.Range[c + l.ToString()].Value);
                            if (protection.Exists(v => v == test.Range[c + l.ToString()].Value)) continue;
                            protection.Add(test.Range[c + l.ToString()].Value);
                        }
                    }
                    for (int l = p2; l < p3 - 6; l += 2)
                    {
                        foreach (string c in col)
                        {
                            if (test.Range[c + l.ToString()].Value == null || test.Range[c + l.ToString()].Value.Replace("\n", "").Replace("\t", "").Replace(" ", "") == string.Empty) continue;
                            if (protection.Exists(v => v == test.Range[c + l.ToString()].Value)) continue;
                            protection.Add(test.Range[c + l.ToString()].Value);
                        }
                    }
                    for (int l = p3; l < p3 + 2 * 7; l += 2)
                    {
                        foreach (string c in col)
                        {
                            if (test.Range[c + l.ToString()].Value == null) continue;
                            if (protection.Exists(v => v == test.Range[c + l.ToString()].Value)) continue;
                            Console.WriteLine(test.Range[c + l.ToString()].Value);
                            protection.Add(test.Range[c + l.ToString()].Value);
                        }
                    }
                    string pl = string.Empty;
                    foreach (string p in protection)
                    {
                        pl += p + "\n";
                    }
                    resultxls.Range["C" + i.ToString()].Value = pl;
                    i++;
                }
                
                result.SaveToFile(filename.Replace(".xlsx", "_result.xlsx").Replace(".xls","_result.xlsx").Replace("_result.xlsxx",".xlsx"), FileFormat.Version2013);
                return true;
            }
            catch(Exception err)
            {
                Console.WriteLine(err.ToString());
                return false;
            }
        }
    }

    
}
