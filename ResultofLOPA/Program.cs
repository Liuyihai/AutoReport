using System;
using System.Collections.Generic;
using Spire.Xls;

namespace ResultofLOPA
{
    class Program
    {
        static void Main(string[] args)
        {
            while(true)
            {
                Console.WriteLine("请输入文件路径或将文件拖入（输入exit退出程序）：");
                string filename = Console.ReadLine();
                GetResultFormLOPA(filename.Replace("\"",""));
            }
        }

        private static void GetResultFormLOPA(string filename)
        {
            Workbook LOPAxls = new Workbook();
            LOPAxls.LoadFromFile(filename);
            Worksheet result = LOPAxls.Worksheets["结论"];
            if(result == null)
            {
                result = LOPAxls.Worksheets.Add("结论");
            }
            
            
            result.Range["A1"].Value = "序号";
            result.Range["B1"].Value = "预设SIF编号(SIF-P)";
            result.Range["C1"].Value = "偏差";
            result.Range["D1"].Value = "后果描述";
            result.Range["E1"].Value = "SIF 需求的要求时失效概率 (PFD)";
            result.Range["G1"].Value = "SIF 需求的SIL等级";
            result.Range["H1"].Value = "SIL定级";
            result.Merge(result.Range["E1"], result.Range["F1"]);
            
            Worksheet list = LOPAxls.Worksheets["SIF list"];
            int i = 2,j = 0;
            foreach(var r in list.Rows)
            {
                if (j == 0 || r.Cells[1].Value == string.Empty)
                {
                    j++;
                    continue;
                }
                Worksheet test = LOPAxls.Worksheets[r.Cells[1].Value];
                if (test == null)
                {
                    test = LOPAxls.Worksheets[r.Cells[6].Value.Replace("参考", "")];
                    result.Merge(result.Range["A" + i.ToString()], result.Range["A" + (i + 2).ToString()]);
                    result.Merge(result.Range["B" + i.ToString()], result.Range["B" + (i + 2).ToString()]);
                    result.Merge(result.Range["C" + i.ToString()], result.Range["C" + (i + 2).ToString()]);
                    result.Merge(result.Range["D" + i.ToString()], result.Range["D" + (i + 2).ToString()]);
                    result.Merge(result.Range["H" + i.ToString()], result.Range["H" + (i + 2).ToString()]);

                    result.Range["A" + i.ToString()].Value = r.Columns[0].Value;
                    result.Range["B" + i.ToString()].Value = r.Columns[1].Value;
                    result.Range["C" + i.ToString()].Value = test.Range["B5"].Value;
                    result.Range["D" + i.ToString()].Value = test.Range["C5"].Value;
                    result.Range["E" + i.ToString()].Value = "人员安全";
                    result.Range["E" + (i + 1).ToString()].Value = "环境影响";
                    result.Range["E" + (i + 2).ToString()].Value = "财务风险";

                    result.Range["F" + i.ToString()].Value = test.Range["Q" + test.Range["A1"].FormulaValue].FormulaValue.ToString();
                    result.Range["F" + (i + 1).ToString()].Value = test.Range["Q" + test.Range["B1"].FormulaValue].FormulaValue.ToString();
                    result.Range["F" + (i + 2).ToString()].Value = test.Range["Q" + test.Range["C1"].FormulaValue].FormulaValue.ToString();
                    result.Range["G" + i.ToString()].Value = test.Range["R" + test.Range["A1"].FormulaValue].FormulaValue.ToString();
                    result.Range["G" + (i + 1).ToString()].Value = test.Range["R" + test.Range["B1"].FormulaValue].FormulaValue.ToString();
                    result.Range["G" + (i + 2).ToString()].Value = test.Range["R" + test.Range["C1"].FormulaValue].FormulaValue.ToString();
                    string v1 = result.Range["G" + i.ToString()].Value.Replace("SIL ", "").Replace("SIL", "").Replace("NO", "0").Replace("a", "0.5");
                    string v2 = result.Range["G" + (i + 1).ToString()].Value.Replace("SIL ", "").Replace("SIL", "").Replace("NO", "0").Replace("a", "0.5");
                    string v3 = result.Range["G" + (i + 2).ToString()].Value.Replace("SIL ", "").Replace("SIL", "").Replace("NO", "0").Replace("a", "0.5");
                    double max = Math.Max(double.Parse(v1), Math.Max(double.Parse(v2), double.Parse(v3)));
                    string SIL = string.Empty;
                    if (max.ToString() == v1)
                        SIL = result.Range["G" + i.ToString()].Value;
                    else
                    if (max.ToString() == v2)
                    {
                        SIL = result.Range["G" + (i + 1).ToString()].Value;
                    }
                    else
                        SIL = result.Range["G" + (i + 2).ToString()].Value;
                    result.Range["H" + i.ToString()].Value = SIL;
                    result.Range["I" + i.ToString()].Value = r.Cells[6].Value;
                    i += 3;
                }
                else
                {
                    result.Merge(result.Range["A" + i.ToString()], result.Range["A" + (i + 2).ToString()]);
                    result.Merge(result.Range["B" + i.ToString()], result.Range["B" + (i + 2).ToString()]);
                    result.Merge(result.Range["C" + i.ToString()], result.Range["C" + (i + 2).ToString()]);
                    result.Merge(result.Range["D" + i.ToString()], result.Range["D" + (i + 2).ToString()]);
                    result.Merge(result.Range["H" + i.ToString()], result.Range["H" + (i + 2).ToString()]);

                    result.Range["A" + i.ToString()].Value = r.Columns[0].Value;
                    result.Range["B" + i.ToString()].Value = r.Columns[1].Value;
                    result.Range["C" + i.ToString()].Value = test.Range["B5"].Value;
                    result.Range["D" + i.ToString()].Value = test.Range["C5"].Value;
                    result.Range["E" + i.ToString()].Value = "人员安全";
                    result.Range["E" + (i + 1).ToString()].Value = "环境影响";
                    result.Range["E" + (i + 2).ToString()].Value = "财务风险";

                    result.Range["F" + i.ToString()].Value = test.Range["Q" + test.Range["A1"].FormulaValue].FormulaValue.ToString();
                    result.Range["F" + (i + 1).ToString()].Value = test.Range["Q" + test.Range["B1"].FormulaValue].FormulaValue.ToString();
                    result.Range["F" + (i + 2).ToString()].Value = test.Range["Q" + test.Range["C1"].FormulaValue].FormulaValue.ToString();
                    result.Range["G" + i.ToString()].Value = test.Range["R" + test.Range["A1"].FormulaValue].FormulaValue.ToString();
                    result.Range["G" + (i + 1).ToString()].Value = test.Range["R" + test.Range["B1"].FormulaValue].FormulaValue.ToString();
                    result.Range["G" + (i + 2).ToString()].Value = test.Range["R" + test.Range["C1"].FormulaValue].FormulaValue.ToString();
                    string v1 = result.Range["G" + i.ToString()].Value.Replace("SIL ", "").Replace("SIL", "").Replace("NO", "0").Replace("a", "0.5");
                    string v2 = result.Range["G" + (i + 1).ToString()].Value.Replace("SIL ", "").Replace("SIL", "").Replace("NO", "0").Replace("a", "0.5");
                    string v3 = result.Range["G" + (i + 2).ToString()].Value.Replace("SIL ", "").Replace("SIL", "").Replace("NO", "0").Replace("a", "0.5");
                    double max = Math.Max(double.Parse(v1), Math.Max(double.Parse(v2), double.Parse(v3)));
                    string SIL = string.Empty;
                    if (max.ToString() == v1)
                        SIL = result.Range["G" + i.ToString()].Value;
                    else
                    if (max.ToString() == v2)
                    {
                        SIL = result.Range["G" + (i + 1).ToString()].Value;
                    }
                    else
                        SIL = result.Range["G" + (i + 2).ToString()].Value;
                    result.Range["H" + i.ToString()].Value = SIL;

                    //part 1
                    int p1 = int.Parse(test.Range["A1"].FormulaValue.ToString());
                    int p2 = int.Parse(test.Range["B1"].FormulaValue.ToString());
                    int p3 = int.Parse(test.Range["C1"].FormulaValue.ToString());
                    string[] col = new string[] { "I", "J", "K", "L", "M" };
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
                        pl += p + ";";
                    }
                    result.Range["I" + i.ToString()].Value = pl;
                    i += 3;
                }
                
            }

            LOPAxls.Save();
            Console.WriteLine("信息提取完成，請在源文件目錄查看提取結果！\n\n");
        }
    }
}
