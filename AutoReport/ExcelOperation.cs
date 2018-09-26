using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Spire.Xls;
using Spire.Doc.Documents;
using Spire.Doc;
using Spire.Doc.Fields;
using System.Drawing;

namespace AutoReport
{
    public class ExcelOperation
    {
        public ExcelOperation(String filename)
        {
            this.filename = filename;
        }

        String filename = string.Empty;

        public Document Excel2Docx()
        {
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(filename);
            

            Document doc = new Document();

            foreach(Worksheet sheet in workbook.Worksheets)
            {
                if (sheet == workbook.Worksheets[0])
                    continue;
                else
                {
                    String headText = sheet.GetText(1, 1) +"  " + sheet.GetText(2, 1);

                    Paragraph paragraph = doc.AddSection().AddParagraph();
                    paragraph.AppendText(headText);
                    paragraph.ApplyStyle(BuiltinStyle.Heading2);

                    Section section = doc.AddSection();
                    Table table = section.AddTable(true);
                    table.ResetCells(10, 6);
                    table.ApplyHorizontalMerge(0, 0, 2);
                    table.ApplyHorizontalMerge(0, 3, 5);
                    table.ApplyHorizontalMerge(1, 0, 2);
                    table.ApplyHorizontalMerge(5, 0, 2);
                    table.ApplyHorizontalMerge(6, 0, 2);
                    table.ApplyHorizontalMerge(1, 3, 5);
                    table.ApplyHorizontalMerge(2, 3, 5);
                    table.ApplyHorizontalMerge(3, 3, 5);
                    table.ApplyHorizontalMerge(4, 3, 5);
                    table.ApplyHorizontalMerge(5, 3, 5);
                    table.ApplyHorizontalMerge(6, 3, 5);
                    table.ApplyVerticalMerge(2, 2, 4);
                    table.ApplyHorizontalMerge(7, 0, 1);
                    table.ApplyHorizontalMerge(8, 0, 1);
                    table.ApplyHorizontalMerge(9, 0, 1);
                    table.ApplyVerticalMerge(0, 7, 9);
                    table.ApplyVerticalMerge(4, 7, 9);
                    table.ApplyVerticalMerge(5, 7, 9);

                    //前两行
                    TextRange range = table[0, 0].AddParagraph().AppendText("SIF描述");
                    range = table[0, 1].AddParagraph().AppendText(sheet.GetText(3, 1));
                    range = table[1, 0].AddParagraph().AppendText("事件后果");
                    range = table[1, 1].AddParagraph().AppendText(sheet.GetText(4, 1));

                    //危害程度
                    range = table[2, 0].AddParagraph().AppendText("危害程度");
                    range = table[2, 1].AddParagraph().AppendText("人员");
                    range = table[3, 1].AddParagraph().AppendText("环境");
                    range = table[4, 1].AddParagraph().AppendText("财产");
                    String str = sheet.GetText(41, 1).Replace("：", "").Replace("0", "").Replace("N", "").Replace("L", "").Replace("M", "").Replace("H", "").Replace("E", "");
                    range = table[2, 2].AddParagraph().AppendText(str);
                    str = sheet.GetText(42,1).Replace("：", "").Replace("0", "").Replace("N", "").Replace("L", "").Replace("M", "").Replace("H", "").Replace("E", "");
                    range = table[3, 2].AddParagraph().AppendText(str);
                    str = sheet.GetText(43,1).Replace("：", "").Replace("0", "").Replace("N", "").Replace("L", "").Replace("M", "").Replace("H", "").Replace("E", "");
                    range = table[4, 2].AddParagraph().AppendText(str);

                    //保护层&减缓措施
                    range = table[5, 0].AddParagraph().AppendText("独立保护层");
                    str = sheet.GetText(45, 1);
                    range = table[5, 1].AddParagraph().AppendText(str);
                    range = table[6, 0].AddParagraph().AppendText("减缓措施");
                    str = sheet.GetText(46, 1);
                    range = table[6, 1].AddParagraph().AppendText(str);

                    //SIL分项定级
                    range = table[7, 0].AddParagraph().AppendText("SIL分项定级");
                    range = table[7, 1].AddParagraph().AppendText("人员");
                    range = table[8, 1].AddParagraph().AppendText("环境");
                    range = table[9, 1].AddParagraph().AppendText("财产");
                    //人员
                    List<CellRange> cells = new List<CellRange>();
                    cells.Add(sheet.Rows[29]);
                    cells.Add(sheet.Rows[32]);
                    cells.Add(sheet.Rows[35]);
                    cells.Add(sheet.Rows[38]);
                    foreach (CellRange cell in cells)
                        if (cell.Style.Color == Color.Red)
                        {
                            str = cell.Text;
                            break;
                        }
                    range = table[7, 2].AddParagraph().AppendText(str);
                    //环境
                    cells.Clear();
                    cells.Add(sheet.Rows[30]);
                    cells.Add(sheet.Rows[33]);
                    cells.Add(sheet.Rows[36]);
                    cells.Add(sheet.Rows[39]);
                    foreach (CellRange cell in cells)
                        if (cell.Style.Color == Color.Red)
                        {
                            str = cell.Text;
                            break;
                        }
                    range = table[8, 2].AddParagraph().AppendText(str);
                    //财产
                    cells.Clear();
                    cells.Add(sheet.Rows[30]);
                    cells.Add(sheet.Rows[33]);
                    cells.Add(sheet.Rows[36]);
                    cells.Add(sheet.Rows[39]);
                    foreach (CellRange cell in cells)
                        if (cell.Style.Color == Color.Red)
                        {
                            str = cell.Text;
                            break;
                        }
                    range = table[9, 2].AddParagraph().AppendText(str);
                    //最终定级
                    range = table[7, 3].AddParagraph().AppendText("SIL定级");
                    range = table[7, 4].AddParagraph().AppendText(sheet.GetText(43, 7));

                    //指定表格字体及大小
                    ParagraphStyle pStyle = new ParagraphStyle(doc) { Name = "ParagraphStyle" };
                    pStyle.CharacterFormat.FontName = "宋体";
                    pStyle.CharacterFormat.FontSize = 12;
                    doc.Styles.Add(pStyle);

                    foreach(TableRow row in table.Rows)
                    {
                        foreach(TableCell cell in row.Cells)
                        {
                            foreach (Paragraph pg in cell.Paragraphs)
                            {
                                pg.ApplyStyle(pStyle.Name);
                            }
                        }
                        
                    }

                    DefaultTableStyle style = DefaultTableStyle.TableGrid;
                    table.ApplyStyle(style);
                }
            }

            doc.SaveToFile(@"\Data\SILlevel\temp.docx");
            return doc;
        }
    }
}
