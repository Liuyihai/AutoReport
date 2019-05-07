using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Spire.Xls;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using System.Drawing;
using System.Text.RegularExpressions;

namespace Test4ExcelOperation
{
    public class ExcelOperation
    {
        public void Excel2Docx(string filename)
        {
            filename = filename.Replace("\"", "");
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(filename);

            //Spire.Xls.Core.IConditionalFormat;
            Document doc = new Document();
            Section section = doc.AddSection();
            //指定表格字体及大小
            ParagraphStyle sty = new ParagraphStyle(doc)
            {
                Name = "fortable"
            };
            sty.CharacterFormat.FontName = "宋体";
            sty.CharacterFormat.FontSize = 12;
            sty.CharacterFormat.Italic = false;
            doc.Styles.Add(sty);
            Table lead = section.AddTable();
            lead.ResetCells(1, 4);
            lead.Rows[0].Cells[0].AddParagraph().AppendText("No");
            lead.Rows[0].Cells[1].AddParagraph().AppendText("SIF No");
            lead.Rows[0].Cells[2].AddParagraph().AppendText("SIF Name");
            lead.Rows[0].Cells[3].AddParagraph().AppendText("SIL");
            DefaultTableStyle style = DefaultTableStyle.TableGrid;
            lead.ApplyStyle(style);
            int No = 0;
            int count_NA = 0;
            int count_SIL1 = 0;
            int count_SIL2 = 0;
            int count_SIL3 = 0;
            int count_SIL4 = 0;

            Paragraph p = new Paragraph(doc);
            p.ApplyStyle(BuiltinStyle.Heading2);
            ParagraphStyle header = p.GetStyle();
            header.Name = "Header";
            header.CharacterFormat.FontName = "宋体";
            header.CharacterFormat.Italic = false;
            doc.Styles.Add(header);


            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Regex regex = new Regex(@"SIF List");
                
                Console.WriteLine(sheet.Name);
                if (sheet == workbook.Worksheets[0] || sheet.Name == "SIL decision matrix")
                    continue;
                else if (regex.Match(sheet.Name).Success) continue;
                else
                {
                    No++;
                    TableRow hr = new TableRow(doc);
                    hr.AddCell();
                    hr.Cells[hr.Cells.Count - 1].AddParagraph().AppendText(No.ToString());
                    hr.AddCell();
                    hr.Cells[hr.Cells.Count - 1].AddParagraph().AppendText(sheet.Range["B2"].FormulaValue.ToString());
                    hr.AddCell();
                    Paragraph cellcontent = hr.Cells[hr.Cells.Count - 1].AddParagraph();
                    cellcontent.AppendText(sheet.Range["B3"].FormulaValue.ToString());
                    cellcontent.ApplyStyle(sty.Name);
                    hr.AddCell();
                    Console.WriteLine(sheet.Range["H3"].FormulaValue);
                    string cell4 = sheet.Range["H3"].FormulaValue.ToString();
                    hr.Cells[hr.Cells.Count - 1].AddParagraph().AppendText(cell4);

                    lead.Rows.Add(hr);
                    String headText = sheet.Range["B2"].FormulaValue + "  " + sheet.Range["B3"].FormulaValue;
                    Console.WriteLine("", sheet.Range["B2"].FormulaValue, sheet.Range["B3"].FormulaValue);
                    Paragraph paragraph = section.AddParagraph();
                    paragraph.AppendText(headText);
                    paragraph.ApplyStyle(header.Name);
                    

                    Table table = section.AddTable(true);
                    table.ResetCells(10, 6);


                    //前两行
                    TableRow row = table.Rows[0];
                    TextRange range = row.Cells[0].AddParagraph().AppendText("SIF描述");
                    range = row.Cells[3].AddParagraph().AppendText(sheet.Range["B4"].Text);
                    row = table.Rows[1];
                    range = row.Cells[0].AddParagraph().AppendText("事件后果");
                    range = row.Cells[3].AddParagraph().AppendText(sheet.Range["B5"].Text);

                    //危害程度
                    range = table.Rows[2].Cells[0].AddParagraph().AppendText("危害程度");
                    range = table.Rows[2].Cells[1].AddParagraph().AppendText("人员");
                    range = table.Rows[3].Cells[1].AddParagraph().AppendText("环境");
                    range = table.Rows[4].Cells[1].AddParagraph().AppendText("财产");
                    String strg = sheet.GetText(42, 2);
                    if (strg != null)
                    {
                        strg = strg.Replace("：", "").Replace("0", "").Replace("N", "").Replace("L", "").Replace("M", "").Replace("H", "").Replace("E", "");

                        range = table[2, 3].AddParagraph().AppendText(strg);
                    }
                    strg = sheet.GetText(43, 2);
                    if (strg != null)
                    {
                        strg = strg.Replace("：", "").Replace("0", "").Replace("N", "").Replace("L", "").Replace("M", "").Replace("H", "").Replace("E", "");

                        range = table.Rows[3].Cells[3].AddParagraph().AppendText(strg);
                    }
                    strg = sheet.GetText(44, 2);
                    if (strg != null)
                    {
                        strg = strg.Replace("：", "").Replace("0", "").Replace("N", "").Replace("L", "").Replace("M", "").Replace("H", "").Replace("E", "");
                        range = table[4, 3].AddParagraph().AppendText(strg);
                    }


                    //保护层&减缓措施
                    range = table[5, 0].AddParagraph().AppendText("独立保护层");
                    strg = sheet.Range["D22"].Text + sheet.Range["D23"].Text + sheet.Range["D24"].Text;
                    range = table[5, 3].AddParagraph().AppendText(strg);
                    range = table[6, 0].AddParagraph().AppendText("减缓措施");
                    strg = sheet.Range["D12"].Text + sheet.Range["D13"].Text;
                    range = table[6, 3].AddParagraph().AppendText(strg);

                    //SIL分项定级
                    range = table[7, 0].AddParagraph().AppendText("SIL分项定级");
                    range = table[7, 2].AddParagraph().AppendText("人员");
                    range = table[8, 2].AddParagraph().AppendText("环境");
                    range = table[9, 2].AddParagraph().AppendText("财产");

                    strg = string.Empty;
                    string[] head = new string[6] { "B", "C", "D", "E", "F", "G" };
                    int[] num = new int[4] { 30, 33, 36, 39 };
                    int[] cols = new int[6] { 1, 2, 3, 4, 5, 6 };
                    foreach (int i in num)
                    {
                        //人员
                        foreach (string h in head)
                        {
                            Console.WriteLine(sheet.Range[h + i.ToString()].Style.Color.ToString() + "\t" + h + i.ToString());
                            if (sheet.Range[h + i.ToString()].Style.Color.ToArgb() == Color.Red.ToArgb())
                            {
                                strg = sheet.Range[h + i.ToString()].Text;
                                if (strg == null || strg.Replace(" ", "") == string.Empty)
                                {
                                    range = table[7, 3].AddParagraph().AppendText("NA");
                                }
                                else
                                    range = table[7, 3].AddParagraph().AppendText(strg);
                                break;
                            }
                        }
                        if (strg != string.Empty)
                        {
                            //环境
                            int l = i + 1;
                            foreach (string h in head)
                            {
                                if (sheet.Range[h + l.ToString()].Style.Color.ToArgb() == Color.Red.ToArgb())
                                {
                                    strg = sheet.Range[h + l.ToString()].Text;
                                    if (strg == null || strg.Replace(" ", "") == string.Empty)
                                    {
                                        range = table[8, 3].AddParagraph().AppendText("NA");
                                    }
                                    else
                                        range = table[8, 3].AddParagraph().AppendText(strg);
                                    break;
                                }
                            }
                            //财产

                            l++;
                            foreach (string h in head)
                            {
                                if (sheet.Range[h + l.ToString()].Style.Color.ToArgb() == Color.Red.ToArgb())
                                {
                                    strg = sheet.Range[h + l.ToString()].Text;
                                    if (strg == null || strg.Replace(" ", "") == string.Empty)
                                    {
                                        range = table[9, 3].AddParagraph().AppendText("NA");
                                    }
                                    else
                                        range = table[9, 3].AddParagraph().AppendText(strg);
                                    break;
                                }
                            }
                        }
                    }


                    //最终定级
                    range = table[7, 4].AddParagraph().AppendText("SIL定级");
                    range = table[7, 5].AddParagraph().AppendText(sheet.GetText(44, 8));
                    if (sheet.GetText(44, 8) == "NA")
                        count_NA++;
                    else if (sheet.GetText(44, 8) == "SIL 1")
                        count_SIL1++;
                    else if (sheet.GetText(44, 8) == "SIL 2")
                        count_SIL2++;
                    else if (sheet.GetText(44, 8) == "SIL 3")
                        count_SIL3++;
                    else count_SIL4++;
                        

                    int[] aa = new int[7] { 0, 1, 2, 3, 4, 5, 6 };
                    foreach (TableRow rowl in table.Rows)
                    {
                        foreach (TableCell cell in rowl.Cells)
                        {
                            cell.CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                            foreach (Paragraph pg in cell.Paragraphs)
                            {
                                pg.ApplyStyle(sty.Name);
                                pg.Format.HorizontalAlignment = HorizontalAlignment.Center;
                            }

                            foreach (int a in aa)
                            {
                                if (rowl == table.Rows[a])
                                {
                                    if (cell == rowl.Cells[3])
                                    {
                                        cell.CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                                        foreach (Paragraph pg in cell.Paragraphs)
                                        {
                                            pg.ApplyStyle(sty.Name);
                                            pg.Format.HorizontalAlignment = HorizontalAlignment.Left;
                                        }

                                    }
                                    else
                                    {
                                        cell.CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                                        foreach (Paragraph pg in cell.Paragraphs)
                                        {
                                            pg.ApplyStyle(sty.Name);
                                            pg.Format.HorizontalAlignment = HorizontalAlignment.Center;
                                        }
                                    }

                                }
                                else
                                    continue;
                            }


                        }

                    }

                    table.ApplyHorizontalMerge(0, 0, 2);
                    table.ApplyHorizontalMerge(0, 3, 5);
                    table.ApplyHorizontalMerge(1, 0, 2);
                    table.ApplyHorizontalMerge(5, 0, 2);
                    table.ApplyHorizontalMerge(6, 0, 2);
                    table.ApplyHorizontalMerge(1, 3, 5);
                    table.ApplyHorizontalMerge(2, 1, 2);
                    table.ApplyHorizontalMerge(3, 1, 2);
                    table.ApplyHorizontalMerge(4, 1, 2);
                    table.ApplyHorizontalMerge(2, 3, 5);
                    table.ApplyHorizontalMerge(3, 3, 5);
                    table.ApplyHorizontalMerge(4, 3, 5);
                    table.ApplyHorizontalMerge(5, 3, 5);
                    table.ApplyHorizontalMerge(6, 3, 5);
                    table.ApplyVerticalMerge(0, 2, 4);
                    table.ApplyHorizontalMerge(7, 0, 1);
                    table.ApplyHorizontalMerge(8, 0, 1);
                    table.ApplyHorizontalMerge(9, 0, 1);
                    table.ApplyVerticalMerge(0, 7, 9);
                    table.ApplyVerticalMerge(4, 7, 9);
                    table.ApplyVerticalMerge(5, 7, 9);

                    table.ApplyStyle(style);
                }
            }

            Paragraph pgg = doc.LastSection.AddParagraph();
            string str = "NA等级联锁数量：" + count_NA + "\nSIL 1等级数量：" + count_SIL1 + "\nSIL 2等级数量：" +
                count_SIL2 + "\nSIL 3等级联锁数量：" + count_SIL3 + "\nSIL 4等级联锁数量：" + count_SIL4;
            pgg.AppendText(str);

            doc.SaveToFile(filename.Replace(".xlsm",".docx"), Spire.Doc.FileFormat.Docx2013);
            doc.Close();
        }
    }
}
