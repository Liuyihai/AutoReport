using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AutoReport
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        Report_Type type = Report_Type.Risk_SILlevel;


        private void 风险分析与SILToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_Click(object sender, EventArgs e)
        {
            textBox1.Text = string.Empty;
        }

        private void textBox2_Click(object sender, EventArgs e)
        {
            textBox2.Text = string.Empty;
        }

        private void textBox3_Click(object sender, EventArgs e)
        {
            textBox3.Text = string.Empty;
        }

        private void textBox4_Click(object sender, EventArgs e)
        {
            textBox4.Text = string.Empty;

        }

        private void textBox5_Click(object sender, EventArgs e)
        {
                       
        }

        private void textBox6_Click(object sender, EventArgs e)
        {
            textBox6.Text = string.Empty;

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
            }
            Operation op = GetOperation();
            op.DocMerge(filename.ToString(),this);
            MessageBox.Show("报告已生成完毕。");
        }

        private Operation GetOperation()
        {
            Operation op = null;
            if (type == Report_Type.Risk_SILlevel)
                op = new Risk_SILlevel();
            if(type == Report_Type.SIL_analysis)
                op = new SIL_analysis();
            if (type == Report_Type.SIL_validate)
                op = new SIL_validate();
            if (type == Report_Type.MTTR_analysis)
                op = new MTTR_analysis();

            return op;   
        }
    }
}
