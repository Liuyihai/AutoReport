using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Spire.Doc;
using System.Windows.Forms;
using Spire.Doc.Documents;

namespace AutoReport
{
    class Risk_SILlevel : Operation
    {
        public Risk_SILlevel()
        {

        }
        
        /// <summary>
        /// 实现父类的合并文档抽象方法
        /// </summary>
        public override void DocMerge(String filename,Form1 form)
        {
            Document doc1 = new Document();
            doc1.LoadFromFile(System.Environment.CurrentDirectory + @"\Data\SILlevel\doc1.docx");

            Document doc2 = new Document();
            doc2.LoadFromFile(System.Environment.CurrentDirectory + @"\Data\SILlevel\doc2.docx");

            Document doc3 = new Document();

            Paragraph paragraph = doc3.AddSection().AddParagraph();
            paragraph.AppendText("\n");
            paragraph.AppendText("\n");
            paragraph.AppendText("\n");
            paragraph.AppendText(form.textBox8.Text + "\n");
            paragraph.AppendText(form.textBox1.Text + "\n");
            paragraph.AppendText("\n");
            paragraph.AppendText(form.textBox2.Text + "\n");
            paragraph.AppendText("风险分析与SIL定级报告");
            paragraph.AppendText("\n");
            paragraph.AppendText("\n");
            paragraph.AppendText("\n");
            paragraph.AppendText("\n");
            paragraph.AppendText("\n");
            paragraph.AppendText("\n");
            paragraph.AppendText("\n");
            paragraph.AppendText("\n");
            paragraph.AppendText("\n");
            ParagraphStyle style = new ParagraphStyle(doc3)
            {
                Name = "firstpage"
            };
            style.CharacterFormat.Bold = true;
            style.CharacterFormat.FontSize = 20;
            style.CharacterFormat.FontName = "宋体";
            doc3.Styles.Add(style);

            paragraph.ApplyStyle(style.Name);

            

            foreach (Section sec in doc1.Sections)
            {
                foreach (DocumentObject obj in sec.Body.ChildObjects)
                {
                    doc3.LastSection.Body.ChildObjects.Add(obj.Clone());
                }
            }

            foreach (Section sec in doc2.Sections)
            {
                foreach (DocumentObject obj in sec.Body.ChildObjects)
                {
                    doc3.LastSection.Body.ChildObjects.Add(obj.Clone());
                }
            }

            doc3.SaveToFile(filename, FileFormat.Docx2013);
            doc1.Close();
            doc2.Close();
            doc3.Close();
        }
    }
}
