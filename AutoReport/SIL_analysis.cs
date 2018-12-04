using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Spire.Doc;
using Spire.Xls;

namespace AutoReport
{
    class SIL_analysis : Operation
    {
        public override void DocMerge(String filename,Form1 form)
        {
            Document saveDocFile = new Document();

            //合并doc1.docx
            Document model = new Document();
            model.LoadFromFile(@"./Data/SILlevel/doc1.docx");
            foreach (Section sec in model.Sections)
            {
                Section section = saveDocFile.AddSection();
                foreach (DocumentObject obj in sec.Body.ChildObjects)
                {
                    saveDocFile.LastSection.Body.ChildObjects.Add(obj.Clone());
                }
            }
            model.Close();
            //导入表格
            ExcelOperation excelOperation = new ExcelOperation(form.textBox5.Text);
            Document excel = excelOperation.Excel2Docx();
            foreach (Section sec in excel.Sections)
            {
                Section section = saveDocFile.AddSection();
                foreach (DocumentObject obj in sec.Body.ChildObjects)
                {
                    saveDocFile.LastSection.Body.ChildObjects.Add(obj.Clone());
                }
            }
            excel.Close();
            //合并doc2.docx
            model.LoadFromFile(@"./Data/SILlevel/doc2.docx");
            foreach (Section sec in model.Sections)
            {
                Section section = saveDocFile.LastSection;
                foreach (DocumentObject obj in sec.Body.ChildObjects)
                {
                    saveDocFile.LastSection.Body.ChildObjects.Add(obj.Clone());
                }
            }
            model.Close();
            //添加SIL验算信息
            Document report = new Document();
            report.LoadFromFile(form.textBox3.Text);
            //单个表格模板
            Table modelTab = new Table(saveDocFile);
            modelTab.ResetCells(18, 4);
            modelTab.Rows[0].Cells[0].AddParagraph().AppendText("安全仪表功能参数");
            modelTab.Rows[1].Cells[0].AddParagraph().AppendText("SIL 目标");
            modelTab.Rows[2].Cells[0].AddParagraph().AppendText("RRF 目标");
            modelTab.Rows[3].Cells[0].AddParagraph().AppendText("SIL 实现");
            modelTab.Rows[4].Cells[0].AddParagraph().AppendText("");

            foreach(Section sec in report.Sections)
            {
                foreach(Table table in sec.Tables)
                {
                    string text = table.Rows[2].Cells[1].Paragraphs.ToString() + table.Rows[3].Cells[1].Paragraphs.ToString();
                    text.Replace("\n", " ");
                    saveDocFile.LastParagraph.AppendText(text);
                    table.Rows[3].Cells[0].Tables[0].Rows[0].Cells[0].AddParagraph().AppendText("");
                }
            }

        }
    }
}
