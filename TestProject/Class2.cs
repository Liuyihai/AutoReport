using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Spire.Xls;

namespace TestProject
{
    class Class2
    {
        public void DataDeal(string path)
        {
            Workbook xls = new Workbook();
            xls.LoadFromFile(path);
            Worksheet sheet = xls.Worksheets[0];
            Workbook result = new Workbook();
            Worksheet resultsheet = result.Worksheets[0];
            List<DataRow> data = new List<DataRow>();
            List<Class1> res = new List<Class1>();
            foreach (var row in sheet.Rows)
            {
                string wellname = row.Cells[0].Value;
                string xch = row.Cells[1].Value;
                if (wellname == "WellName")
                    continue;
                DataRow dr = new DataRow
                {
                    WellNname = wellname,
                    XCH = xch
                };
                Class1 info = new Class1
                {
                    WellName = wellname,
                    XCH = xch,
                    TOP = row.Cells[2].Value,
                    BOT = row.Cells[3].Value
                };
                foreach (var rw in sheet.Rows)
                {
                    string wn = rw.Cells[0].Value;
                    string xh = rw.Cells[1].Value;
                    DataRow drw = new DataRow
                    {
                        WellNname = wn,
                        XCH = xh
                    };
                    if (drw.WellNname == dr.WellNname && drw.XCH == dr.XCH)
                    {
                        if (rw.Cells[2].Value != null && rw.Cells[3].Value != null)
                        {
                            info.TOP = Math.Min(double.Parse(info.TOP), double.Parse(rw.Cells[2].Value)).ToString();
                            info.BOT = Math.Max(double.Parse(info.BOT), double.Parse(rw.Cells[3].Value)).ToString();
                        }
                        else if (rw.Cells[2].Value == null)
                        {
                            info.TOP = string.Empty;
                            if (rw.Cells[3].Value == null)
                            {
                                info.BOT = string.Empty;
                            }
                            else
                            {
                                info.BOT = rw.Cells[3].Value;
                            }
                        }
                        else
                        {
                            info.TOP = rw.Cells[2].Value;
                            info.BOT = string.Empty;
                        }


                    }
                }
                foreach (var m in res)
                {
                    if (m.WellName == dr.WellNname && m.XCH == dr.XCH)
                    {
                        info = null;
                        break;
                    }
                }
                if (info == null)
                    continue;
                res.Add(info);
                Console.WriteLine(info.WellName + "\t" + info.XCH + "\t" + info.TOP + "\t" + info.BOT + "\n");

            }
            int i = 2;
            resultsheet.Range["A1"].Value = "wellName";
            resultsheet.Range["B1"].Value = "XCH";
            resultsheet.Range["C1"].Value = "TOP";
            resultsheet.Range["D1"].Value = "BOT";
            foreach (var n in res)
            {
                resultsheet.Range["A" + i.ToString()].Value = n.WellName;
                resultsheet.Range["B" + i.ToString()].Value = n.XCH;
                resultsheet.Range["C" + i.ToString()].Value = n.TOP;
                resultsheet.Range["D" + i.ToString()].Value = n.BOT;
                i++;
            }

            result.SaveToFile(path.Replace(".xls", "_result.xls"), FileFormat.Version2013);
            Console.WriteLine("提取完成\n");
        }
    }
}
