using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Spire.Doc;
using Spire.Doc.Documents;

namespace GetExcelFormReoprt
{
    class Program
    {
        static void Main(string[] args)
        {
            while(true)
            {
                Console.WriteLine("请输入摘取后的exida报告文件路径或直接将其拖入此窗口：");
                string filename = Console.ReadLine();
                GetExcel(filename);
            }
        }

        private static void GetExcel(string filename)
        {
            Document saveDocFile = new Document();
            //添加SIL验算信息
            Document report = new Document();
            report.LoadFromFile(filename);
            Section section = saveDocFile.AddSection();
            //添加统计表格
            Table total = section.AddTable();
            total.ResetCells(1, 6);
            total.Rows[0].Cells[0].AddParagraph().AppendText("序号");
            total.Rows[0].Cells[0].Width = 20;
            total.Rows[0].Cells[1].AddParagraph().AppendText("SIF编号");
            total.Rows[0].Cells[1].Width = 80;
            total.Rows[0].Cells[2].AddParagraph().AppendText("SIF名称");
            total.Rows[0].Cells[2].Width = 120;
            total.Rows[0].Cells[3].AddParagraph().AppendText("SIL需求");
            total.Rows[0].Cells[3].Width = 60;
            total.Rows[0].Cells[4].AddParagraph().AppendText("SIL实现");
            total.Rows[0].Cells[4].Width = 60;
            total.Rows[0].Cells[5].AddParagraph().AppendText("误停车时间");
            total.Rows[0].Cells[5].Width = 60;
            total.ApplyStyle(DefaultTableStyle.TableGrid);

            ParagraphStyle TitleOfExcel = new ParagraphStyle(saveDocFile);
            TitleOfExcel.Name = "TitleOfExcel";
            TitleOfExcel.CharacterFormat.FontName = "宋体";
            TitleOfExcel.CharacterFormat.FontSize = 12;
            saveDocFile.Styles.Add(TitleOfExcel);

            ParagraphStyle cellstyle = new ParagraphStyle(saveDocFile);
            cellstyle.Name = "CellStyle";
            cellstyle.CharacterFormat.FontName = "宋体";
            cellstyle.CharacterFormat.FontSize = 12;
            saveDocFile.Styles.Add(cellstyle);

            foreach (Section sec in report.Sections)
            {
                foreach (Table table in sec.Tables)
                {
                    total.AddRow(true);
                    total.LastRow.Cells[0].AddParagraph().AppendText((total.Rows.Count - 1).ToString());
                    total.LastRow.Cells[1].AddParagraph().AppendText(table.Rows[2].Cells[1].Paragraphs[0].Text);
                    total.LastRow.Cells[2].AddParagraph().AppendText(table.Rows[3].Cells[1].Paragraphs[0].Text);
                    string value = "SIL " + table.Rows[5].Cells[0].Tables[0].Rows[1].Cells[1].Paragraphs[0].Text;
                    total.LastRow.Cells[3].AddParagraph().AppendText(value);
                    value = "SIL " + table.Rows[5].Cells[0].Tables[0].Rows[4].Cells[1].Paragraphs[0].Text;
                    total.LastRow.Cells[4].AddParagraph().AppendText(value);
                    value = table.Rows[5].Cells[0].Tables[0].Rows[10].Cells[1].Paragraphs[0].Text + "years";
                    total.LastRow.Cells[5].AddParagraph().AppendText(value);

                    string text = table.Rows[2].Cells[1].Paragraphs[0].Text + table.Rows[3].Cells[1].Paragraphs[0].Text;
                    text.Replace("\n", " ");
                    Paragraph pg = section.AddParagraph();
                    pg.AppendText(text);
                    //pg.ApplyStyle(TitleOfExcel.Name);
                    pg.ApplyStyle(BuiltinStyle.Heading2);

                    Table modelTab = saveDocFile.LastSection.AddTable();
                    modelTab.ResetCells(18, 4);
                    foreach (TableRow row in modelTab.Rows)
                    {
                        foreach (TableCell cell in row.Cells)
                        {
                            cell.Width = 95;
                        }
                    }
                    //modelTab.ColumnWidth = new float[4] { 70, 70, 70, 70 };
                    modelTab.Rows[0].Cells[0].AddParagraph().AppendText("安全仪表功能参数");
                    modelTab.ApplyHorizontalMerge(0, 0, 3);
                    modelTab.Rows[1].Cells[0].AddParagraph().AppendText("SIL 目标");
                    modelTab.ApplyHorizontalMerge(1, 0, 1);
                    modelTab.Rows[2].Cells[0].AddParagraph().AppendText("RRF 目标");
                    modelTab.ApplyHorizontalMerge(2, 0, 1);
                    modelTab.Rows[3].Cells[0].AddParagraph().AppendText("SIL 实现");
                    modelTab.ApplyHorizontalMerge(3, 0, 1);
                    modelTab.Rows[4].Cells[0].AddParagraph().AppendText("PFDavg");
                    modelTab.ApplyHorizontalMerge(4, 0, 1);
                    modelTab.Rows[5].Cells[0].AddParagraph().AppendText("SIL (PFDavg)");
                    modelTab.ApplyHorizontalMerge(5, 0, 1);
                    modelTab.Rows[6].Cells[0].AddParagraph().AppendText("SIL 约束");
                    modelTab.ApplyHorizontalMerge(6, 0, 1);
                    modelTab.Rows[7].Cells[0].AddParagraph().AppendText("SIL 能力");
                    modelTab.ApplyHorizontalMerge(7, 0, 1);
                    modelTab.Rows[8].Cells[0].AddParagraph().AppendText("RRF 实现");
                    modelTab.ApplyHorizontalMerge(8, 0, 1);
                    modelTab.Rows[9].Cells[0].AddParagraph().AppendText("MTTFS (年)");
                    modelTab.ApplyHorizontalMerge(9, 0, 1);
                    modelTab.Rows[10].Cells[1].AddParagraph().AppendText("PFDavg");
                    modelTab.Rows[10].Cells[2].AddParagraph().AppendText("MTTFS");
                    modelTab.Rows[10].Cells[3].AddParagraph().AppendText("SILac");
                    modelTab.Rows[11].Cells[0].AddParagraph().AppendText("传感器部分");
                    modelTab.Rows[12].Cells[0].AddParagraph().AppendText("逻辑控制器部分");
                    modelTab.Rows[13].Cells[0].AddParagraph().AppendText("执行器部分");
                    modelTab.Rows[14].Cells[1].AddParagraph().AppendText("MTTR / Hrs");
                    modelTab.Rows[14].Cells[2].AddParagraph().AppendText("PTI / Month");
                    modelTab.Rows[14].Cells[3].AddParagraph().AppendText("PTC / %");
                    modelTab.Rows[15].Cells[0].AddParagraph().AppendText("传感器部分");
                    modelTab.Rows[16].Cells[0].AddParagraph().AppendText("逻辑控制器部分");
                    modelTab.Rows[17].Cells[0].AddParagraph().AppendText("执行器部分");

                    value = table.Rows[5].Cells[0].Tables[0].Rows[1].Cells[1].Paragraphs[0].Text;
                    modelTab.Rows[1].Cells[2].AddParagraph().AppendText(value);
                    modelTab.ApplyHorizontalMerge(1, 2, 3);

                    value = table.Rows[5].Cells[0].Tables[0].Rows[2].Cells[1].Paragraphs[0].Text;
                    modelTab.Rows[2].Cells[2].AddParagraph().AppendText(value);
                    modelTab.ApplyHorizontalMerge(2, 2, 3);

                    value = table.Rows[5].Cells[0].Tables[0].Rows[4].Cells[1].Paragraphs[0].Text;
                    modelTab.Rows[3].Cells[2].AddParagraph().AppendText(value);
                    modelTab.ApplyHorizontalMerge(3, 2, 3);

                    value = table.Rows[5].Cells[0].Tables[0].Rows[5].Cells[1].Paragraphs[0].Text;
                    modelTab.Rows[4].Cells[2].AddParagraph().AppendText(value);
                    modelTab.ApplyHorizontalMerge(4, 2, 3);

                    value = table.Rows[5].Cells[0].Tables[0].Rows[6].Cells[1].Paragraphs[0].Text;
                    modelTab.Rows[5].Cells[2].AddParagraph().AppendText(value);
                    modelTab.ApplyHorizontalMerge(5, 2, 3);

                    value = table.Rows[5].Cells[0].Tables[0].Rows[7].Cells[1].Paragraphs[0].Text;
                    modelTab.Rows[6].Cells[2].AddParagraph().AppendText(value);
                    modelTab.ApplyHorizontalMerge(6, 2, 3);

                    value = table.Rows[5].Cells[0].Tables[0].Rows[8].Cells[1].Paragraphs[0].Text;
                    modelTab.Rows[7].Cells[2].AddParagraph().AppendText(value);
                    modelTab.ApplyHorizontalMerge(7, 2, 3);

                    value = table.Rows[5].Cells[0].Tables[0].Rows[9].Cells[1].Paragraphs[0].Text;
                    modelTab.Rows[8].Cells[2].AddParagraph().AppendText(value);
                    modelTab.ApplyHorizontalMerge(8, 2, 3);

                    value = table.Rows[5].Cells[0].Tables[0].Rows[10].Cells[1].Paragraphs[0].Text;
                    modelTab.Rows[9].Cells[2].AddParagraph().AppendText(value);
                    modelTab.ApplyHorizontalMerge(9, 2, 3);

                    value = table.Rows[5].Cells[0].Tables[1].Rows[2].Cells[1].Paragraphs[0].Text;
                    modelTab.Rows[11].Cells[1].AddParagraph().AppendText(value);

                    value = table.Rows[5].Cells[0].Tables[1].Rows[2].Cells[2].Paragraphs[0].Text;
                    modelTab.Rows[11].Cells[2].AddParagraph().AppendText(value);

                    value = table.Rows[5].Cells[0].Tables[1].Rows[2].Cells[3].Paragraphs[0].Text;
                    modelTab.Rows[11].Cells[3].AddParagraph().AppendText(value);

                    value = table.Rows[5].Cells[0].Tables[1].Rows[3].Cells[1].Paragraphs[0].Text;
                    modelTab.Rows[12].Cells[1].AddParagraph().AppendText(value);

                    value = table.Rows[5].Cells[0].Tables[1].Rows[3].Cells[2].Paragraphs[0].Text;
                    modelTab.Rows[12].Cells[2].AddParagraph().AppendText(value);

                    value = table.Rows[5].Cells[0].Tables[1].Rows[3].Cells[3].Paragraphs[0].Text;
                    modelTab.Rows[12].Cells[3].AddParagraph().AppendText(value);

                    value = table.Rows[5].Cells[0].Tables[1].Rows[4].Cells[1].Paragraphs[0].Text;
                    modelTab.Rows[13].Cells[1].AddParagraph().AppendText(value);

                    value = table.Rows[5].Cells[0].Tables[1].Rows[4].Cells[2].Paragraphs[0].Text;
                    modelTab.Rows[13].Cells[2].AddParagraph().AppendText(value);

                    value = table.Rows[5].Cells[0].Tables[1].Rows[4].Cells[3].Paragraphs[0].Text;
                    modelTab.Rows[13].Cells[3].AddParagraph().AppendText(value);

                    value = "8";
                    modelTab.Rows[15].Cells[1].AddParagraph().AppendText(value);
                    modelTab.Rows[16].Cells[1].AddParagraph().AppendText(value);
                    modelTab.Rows[17].Cells[1].AddParagraph().AppendText(value);

                    value = "36";
                    modelTab.Rows[15].Cells[2].AddParagraph().AppendText(value);
                    modelTab.Rows[16].Cells[2].AddParagraph().AppendText(value);
                    modelTab.Rows[17].Cells[2].AddParagraph().AppendText(value);
                    
                    modelTab.Rows[15].Cells[3].AddParagraph().AppendText("90");
                    modelTab.Rows[16].Cells[3].AddParagraph().AppendText("95");
                    modelTab.Rows[17].Cells[3].AddParagraph().AppendText("85");

                    modelTab.ApplyStyle(DefaultTableStyle.TableGrid);

                    for(int i = 10;i < 18;i++)
                    {
                        foreach(TableCell cell in modelTab.Rows[i].Cells)
                        {
                            if(cell.Paragraphs.Count > 0)
                            {
                                cell.Paragraphs[0].Format.HorizontalAlignment = HorizontalAlignment.Center;
                            }
                        }
                        
                    }

                    for (int i = 1; i < 10; i++)
                    {
                        modelTab.Rows[i].Cells[2].Paragraphs[0].Format.HorizontalAlignment = HorizontalAlignment.Right;
                    }


                    foreach (TableRow row  in modelTab.Rows)
                    {
                        foreach(TableCell cell in row.Cells)
                        {
                            if(cell.Paragraphs.Count > 0)
                            {
                                cell.Paragraphs[0].ApplyStyle(cellstyle.Name);
                            }
                            //cell.Width = 70;
                        }
                    }

                    
                }
            }

            foreach (TableRow row in total.Rows)
            {
                foreach (TableCell cell in row.Cells)
                {
                    if (cell.Paragraphs.Count > 0)
                    {
                        cell.Paragraphs[0].ApplyStyle(cellstyle.Name);
                    }
                }
            }


            report.Close();
            saveDocFile.SaveToFile(filename.Replace(".docx", "_result.docx").Replace(".doc", "_result.docx").Replace("_result.docxx", ".docx"));
            saveDocFile.Close();

        }


    }
}
