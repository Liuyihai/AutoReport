using Spire.Doc;
using Spire.Doc.Documents;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoReport
{
    class MTTR_analysis : Operation
    {
        public override void DocMerge(String filename,Form1 form)
        {
            Document saveDocFile = new Document();

            //合并doc1.docx
            Document model = new Document();
            model.LoadFromFile(@"./Data/MTRanalysis/doc1.docx");
            foreach (Section sec in model.Sections)
            {
                Section section = saveDocFile.AddSection();
                foreach (DocumentObject obj in sec.Body.ChildObjects)
                {
                    saveDocFile.LastSection.Body.ChildObjects.Add(obj.Clone());
                }
            }
            model.Close();
            //添加验算结果
            GetInfoFromReport getInfo = new GetInfoFromReport();
            Document validateResult = getInfo.GetExcel(form.textBox3.Text);
            foreach (Section sec in validateResult.Sections)
            {
                Section section = saveDocFile.AddSection();
                foreach (DocumentObject obj in sec.Body.ChildObjects)
                {
                    saveDocFile.LastSection.Body.ChildObjects.Add(obj.Clone());
                }
            }
            validateResult.Close();
            //合并doc2.docx
            model.LoadFromFile(@"./Data/MTRanalysis/doc2.docx");
            foreach (Section sec in model.Sections)
            {
                Section section = saveDocFile.AddSection();
                foreach (DocumentObject obj in sec.Body.ChildObjects)
                {
                    saveDocFile.LastSection.Body.ChildObjects.Add(obj.Clone());
                }
            }
            model.Close();
            //添加签到表
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
            saveDocFile.Close();
        }
    }
}
