using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Spire.Doc;
using Spire.Doc.Documents;
using System.IO;

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
            GetInfoFromReport getInfo = new GetInfoFromReport();
            Document report = getInfo.GetExcel(form.textBox3.Text);
            foreach(Section sec in report.Sections)
            {
                Section section = saveDocFile.LastSection;
                foreach (DocumentObject obj in sec.Body.ChildObjects)
                {
                    saveDocFile.LastSection.Body.ChildObjects.Add(obj.Clone());
                }
            }
            report.Close();

            //合并doc3.docx
            model.LoadFromFile(@"./Data/SILlevel/doc3.docx");
            foreach (Section sec in model.Sections)
            {
                Section section = saveDocFile.LastSection;
                foreach (DocumentObject obj in sec.Body.ChildObjects)
                {
                    saveDocFile.LastSection.Body.ChildObjects.Add(obj.Clone());
                }
            }
            model.Close();

            //添加会议记录图片
            FileStream fs = File.OpenRead(form.textBox7.Text); //OpenRead
            int filelength = 0;
            filelength = (int)fs.Length; //获得文件长度 
            Byte[] image = new Byte[filelength]; //建立一个字节数组 
            fs.Read(image, 0, filelength);
            Paragraph paragraph = saveDocFile.LastParagraph;
            paragraph.AppendPicture(image);
            fs.Close();

            //字符替换
            string[] replaceText = new string[3]
                { form.textBox8.Text, form.textBox1.Text,  form.textBox2.Text
                };
            ReplaceStr replace = new ReplaceStr();
            saveDocFile = replace.Replace(saveDocFile, replaceText);
            saveDocFile.SaveToFile(filename, FileFormat.Docx2013);
            //saveDocFile.SaveToFile(filename, FileFormat.Docx2013);
            saveDocFile.Close();
        }
    }
}
