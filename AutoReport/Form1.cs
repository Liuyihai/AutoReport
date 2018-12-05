using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Spire.Xls;

namespace AutoReport
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        Report_Type type = Report_Type.Risk_SILlevel;
        //string filepath = 


        private void 风险分析与SILToolStripMenuItem_Click(object sender, EventArgs e)
        {
            type = Report_Type.Risk_SILlevel;
            label4.Text = "当前报告类型：风险分析与SIL定级";
        }

        private void textBox1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "项目完整名称")
                textBox1.Text = string.Empty;
            else
                return;
        }

        private void textBox2_Click(object sender, EventArgs e)
        {
            if (textBox2.Text == "装置完整名称")
                textBox2.Text = string.Empty;
            else
                return;
        }
        
        private void textBox6_Click(object sender, EventArgs e)
        {
            if (textBox6.Text == "请输入保护层或导入保护层文档(每行一个，文档为txt格式，每行一个)")
                textBox6.Text = string.Empty;
            else
                return;
        }

        private void 生成报告ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog
            {
                Filter = "Word Document(*.docx)|*.docx",
                DefaultExt = "Word Document(*.docx)|*.docx"
            };
            object filename = null;

            if (sfd.ShowDialog() == DialogResult.OK)
            {
                filename = sfd.FileName;
                Operation op = GetOperation();
                op.DocMerge(filename.ToString(), this);
                MessageBox.Show("报告已生成完毕。");
            }
            
        }

        private Operation GetOperation()
        {
            Operation op = null;
            switch(type)
            {
                case Report_Type.Risk_SILlevel:
                    {
                        op = new Risk_SILlevel();
                        break;
                    }
                case Report_Type.MTTR_analysis:
                    {
                        op = new MTTR_analysis();
                        break;
                    }
                case Report_Type.SIL_analysis:
                    {
                        op = new SIL_analysis();
                        break;
                    }
                case Report_Type.SIL_validate:
                    {
                        op = new SIL_validate();
                        break;
                    }
            }

            return op;   
        }

        private void textBox8_Click(object sender, EventArgs e)
        {
            if (textBox8.Text == "公司完整名称")
                textBox8.Text = string.Empty;
            else
                return;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog
            {
                Filter = "Microsoft Excel工作表文件|*.xlsm;*.xlsx;*.xls"
            };
            object filename = null;
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                filename = ofd.FileName;
                textBox5.Text = filename.ToString();
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog
            {
                Filter = "图片文件|*.jpg;*.jpeg;*.png;*.bpm;*.tif"
            };
            object filename = null;
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                filename = ofd.FileName;
                textBox7.Text = filename.ToString();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog
            {
                Filter = "文本文档|*.txt"
            };
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                StreamReader sr = new StreamReader(ofd.FileName, Encoding.Default);
                String line;
                textBox6.Text = string.Empty;
                line = sr.ReadLine();
                while (line != null)
                {
                    textBox6.Text += line + "\n";
                    line = sr.ReadLine();
                }

                //textBox6.Text = filename;
            }

        }

        private void sIL分析报告ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            type = Report_Type.SIL_analysis;
            label4.Text = "当前报告类型：SIL分析";
        }

        private void sIL验证报告ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            type = Report_Type.SIL_validate;
            label4.Text = "当前报告类型：SIL验证";
        }

        private void 误停车报告ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            type = Report_Type.MTTR_analysis;
            label4.Text = "当前报告类型：误停车分析";
        }

        private void button4_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog
            {
                Filter = "Microsoft Word文档文件|*.doc;*.docx"
            };
            object filename = null;
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                filename = ofd.FileName;
                textBox3.Text = filename.ToString();
            }
        }

        private void 退出ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog
            {
                Filter = "Microsoft Excel工作表文件|*.xlsm;*.xlsx;*.xls"
            };
            object filename = null;
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                filename = ofd.FileName;
                string protection = Protection(filename.ToString());
                textBox6.Text = protection;
            }
        }

        public string Protection(string filepath)
        {
            Workbook excel = new Workbook();
            excel.LoadFromFile(filepath);
            string protection = string.Empty;
            foreach(Worksheet sheet in excel.Worksheets)
            {
                Regex regex = new Regex(@"SIF List");
                if (sheet == excel.Worksheets[0] || sheet.Name == "SIL decision matrix")
                    continue;
                else if (regex.Match(sheet.Name).Success) continue;
                else
                {
                    if(sheet.Range["B46"].FormulaValue != null)
                        protection += sheet.Range["B46"].FormulaValue.ToString() + "\n";
                }
            }
            return protection;
        }
    }
}
