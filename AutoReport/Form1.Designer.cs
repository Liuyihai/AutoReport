namespace AutoReport
{
    partial class Form1
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.textBox5 = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.textBox6 = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.button2 = new System.Windows.Forms.Button();
            this.textBox7 = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.button3 = new System.Windows.Forms.Button();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.生成报告ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.退出ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.textBox8 = new System.Windows.Forms.TextBox();
            this.label8 = new System.Windows.Forms.Label();
            this.textBox3 = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.button4 = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.报告类型ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.风险分析与SILToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.sIL分析报告ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.sIL验证报告ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.误停车报告ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // textBox1
            // 
            this.textBox1.Font = new System.Drawing.Font("宋体", 11F);
            this.textBox1.Location = new System.Drawing.Point(104, 77);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(451, 24);
            this.textBox1.TabIndex = 0;
            this.textBox1.Text = "项目完整名称";
            this.textBox1.Click += new System.EventHandler(this.textBox1_Click);
            // 
            // textBox2
            // 
            this.textBox2.Font = new System.Drawing.Font("宋体", 11F);
            this.textBox2.Location = new System.Drawing.Point(104, 116);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(451, 24);
            this.textBox2.TabIndex = 0;
            this.textBox2.Text = "装置完整名称";
            this.textBox2.Click += new System.EventHandler(this.textBox2_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("宋体", 11F);
            this.label1.Location = new System.Drawing.Point(26, 83);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(75, 15);
            this.label1.TabIndex = 1;
            this.label1.Text = "项目名称:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("宋体", 11F);
            this.label2.Location = new System.Drawing.Point(26, 122);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(75, 15);
            this.label2.TabIndex = 1;
            this.label2.Text = "装置名称:";
            // 
            // textBox5
            // 
            this.textBox5.Font = new System.Drawing.Font("宋体", 11F);
            this.textBox5.Location = new System.Drawing.Point(104, 156);
            this.textBox5.Name = "textBox5";
            this.textBox5.ReadOnly = true;
            this.textBox5.Size = new System.Drawing.Size(397, 24);
            this.textBox5.TabIndex = 0;
            this.textBox5.Text = "请选择SIL分析表格文件";
            this.textBox5.Click += new System.EventHandler(this.button1_Click);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("宋体", 11F);
            this.label5.Location = new System.Drawing.Point(26, 162);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(75, 15);
            this.label5.TabIndex = 1;
            this.label5.Text = "表格导入:";
            // 
            // button1
            // 
            this.button1.Font = new System.Drawing.Font("宋体", 11F);
            this.button1.Location = new System.Drawing.Point(507, 157);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(48, 22);
            this.button1.TabIndex = 2;
            this.button1.Text = "浏览";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // textBox6
            // 
            this.textBox6.AllowDrop = true;
            this.textBox6.Font = new System.Drawing.Font("宋体", 11F);
            this.textBox6.Location = new System.Drawing.Point(104, 237);
            this.textBox6.Multiline = true;
            this.textBox6.Name = "textBox6";
            this.textBox6.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.textBox6.Size = new System.Drawing.Size(397, 198);
            this.textBox6.TabIndex = 0;
            this.textBox6.Text = "请输入保护层或导入保护层文档(每行一个，文档为txt格式，每行一个)";
            this.textBox6.Click += new System.EventHandler(this.textBox6_Click);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("宋体", 11F);
            this.label6.Location = new System.Drawing.Point(41, 240);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(60, 15);
            this.label6.TabIndex = 1;
            this.label6.Text = "保护层:";
            // 
            // button2
            // 
            this.button2.Font = new System.Drawing.Font("宋体", 11F);
            this.button2.Location = new System.Drawing.Point(507, 238);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(48, 22);
            this.button2.TabIndex = 2;
            this.button2.Text = "浏览";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // textBox7
            // 
            this.textBox7.Font = new System.Drawing.Font("宋体", 11F);
            this.textBox7.Location = new System.Drawing.Point(104, 450);
            this.textBox7.Name = "textBox7";
            this.textBox7.ReadOnly = true;
            this.textBox7.Size = new System.Drawing.Size(397, 24);
            this.textBox7.TabIndex = 0;
            this.textBox7.Text = "请选择会议记录图片";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("宋体", 11F);
            this.label7.Location = new System.Drawing.Point(26, 456);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(75, 15);
            this.label7.TabIndex = 1;
            this.label7.Text = "会议记录:";
            // 
            // button3
            // 
            this.button3.Font = new System.Drawing.Font("宋体", 11F);
            this.button3.Location = new System.Drawing.Point(507, 452);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(48, 22);
            this.button3.TabIndex = 2;
            this.button3.Text = "浏览";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // menuStrip1
            // 
            this.menuStrip1.Font = new System.Drawing.Font("Microsoft YaHei UI", 11F);
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.报告类型ToolStripMenuItem,
            this.生成报告ToolStripMenuItem,
            this.退出ToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(587, 28);
            this.menuStrip1.TabIndex = 3;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // 生成报告ToolStripMenuItem
            // 
            this.生成报告ToolStripMenuItem.Name = "生成报告ToolStripMenuItem";
            this.生成报告ToolStripMenuItem.Size = new System.Drawing.Size(81, 24);
            this.生成报告ToolStripMenuItem.Text = "生成报告";
            this.生成报告ToolStripMenuItem.Click += new System.EventHandler(this.生成报告ToolStripMenuItem_Click);
            // 
            // 退出ToolStripMenuItem
            // 
            this.退出ToolStripMenuItem.Name = "退出ToolStripMenuItem";
            this.退出ToolStripMenuItem.Size = new System.Drawing.Size(51, 24);
            this.退出ToolStripMenuItem.Text = "退出";
            this.退出ToolStripMenuItem.Click += new System.EventHandler(this.退出ToolStripMenuItem_Click);
            // 
            // textBox8
            // 
            this.textBox8.Font = new System.Drawing.Font("宋体", 11F);
            this.textBox8.Location = new System.Drawing.Point(104, 38);
            this.textBox8.Name = "textBox8";
            this.textBox8.Size = new System.Drawing.Size(451, 24);
            this.textBox8.TabIndex = 0;
            this.textBox8.Text = "公司完整名称";
            this.textBox8.Click += new System.EventHandler(this.textBox8_Click);
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Font = new System.Drawing.Font("宋体", 11F);
            this.label8.Location = new System.Drawing.Point(26, 44);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(75, 15);
            this.label8.TabIndex = 1;
            this.label8.Text = "公司名称:";
            // 
            // textBox3
            // 
            this.textBox3.Font = new System.Drawing.Font("宋体", 11F);
            this.textBox3.Location = new System.Drawing.Point(104, 196);
            this.textBox3.Name = "textBox3";
            this.textBox3.ReadOnly = true;
            this.textBox3.Size = new System.Drawing.Size(397, 24);
            this.textBox3.TabIndex = 0;
            this.textBox3.Text = "请选择exida导出的报告文件";
            this.textBox3.Click += new System.EventHandler(this.button1_Click);
            this.textBox3.TextChanged += new System.EventHandler(this.button4_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("宋体", 11F);
            this.label3.Location = new System.Drawing.Point(26, 202);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(75, 15);
            this.label3.TabIndex = 1;
            this.label3.Text = "报告导入:";
            // 
            // button4
            // 
            this.button4.Font = new System.Drawing.Font("宋体", 11F);
            this.button4.Location = new System.Drawing.Point(507, 197);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(48, 22);
            this.button4.TabIndex = 2;
            this.button4.Text = "浏览";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // label4
            // 
            this.label4.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(392, 13);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(191, 12);
            this.label4.TabIndex = 4;
            this.label4.Text = "当前报告类型：风险分析与SIL定级";
            // 
            // 报告类型ToolStripMenuItem
            // 
            this.报告类型ToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.风险分析与SILToolStripMenuItem,
            this.sIL分析报告ToolStripMenuItem,
            this.sIL验证报告ToolStripMenuItem,
            this.误停车报告ToolStripMenuItem});
            this.报告类型ToolStripMenuItem.Name = "报告类型ToolStripMenuItem";
            this.报告类型ToolStripMenuItem.Size = new System.Drawing.Size(81, 24);
            this.报告类型ToolStripMenuItem.Text = "报告类型";
            // 
            // 风险分析与SILToolStripMenuItem
            // 
            this.风险分析与SILToolStripMenuItem.Name = "风险分析与SILToolStripMenuItem";
            this.风险分析与SILToolStripMenuItem.Size = new System.Drawing.Size(234, 24);
            this.风险分析与SILToolStripMenuItem.Text = "风险分析与SIL定级报告";
            // 
            // sIL分析报告ToolStripMenuItem
            // 
            this.sIL分析报告ToolStripMenuItem.Name = "sIL分析报告ToolStripMenuItem";
            this.sIL分析报告ToolStripMenuItem.Size = new System.Drawing.Size(234, 24);
            this.sIL分析报告ToolStripMenuItem.Text = "SIL分析报告";
            // 
            // sIL验证报告ToolStripMenuItem
            // 
            this.sIL验证报告ToolStripMenuItem.Name = "sIL验证报告ToolStripMenuItem";
            this.sIL验证报告ToolStripMenuItem.Size = new System.Drawing.Size(234, 24);
            this.sIL验证报告ToolStripMenuItem.Text = "SIL验证报告";
            // 
            // 误停车报告ToolStripMenuItem
            // 
            this.误停车报告ToolStripMenuItem.Name = "误停车报告ToolStripMenuItem";
            this.误停车报告ToolStripMenuItem.Size = new System.Drawing.Size(234, 24);
            this.误停车报告ToolStripMenuItem.Text = "误停车分析报告";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(587, 490);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button4);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.textBox7);
            this.Controls.Add(this.textBox6);
            this.Controls.Add(this.textBox3);
            this.Controls.Add(this.textBox5);
            this.Controls.Add(this.textBox2);
            this.Controls.Add(this.textBox8);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.menuStrip1);
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "Form1";
            this.Text = "自动报告生成软件 V0.1";
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        public System.Windows.Forms.TextBox textBox1;
        public System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        public System.Windows.Forms.TextBox textBox5;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Button button1;
        public System.Windows.Forms.TextBox textBox6;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Button button2;
        public System.Windows.Forms.TextBox textBox7;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem 退出ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 生成报告ToolStripMenuItem;
        public System.Windows.Forms.TextBox textBox8;
        private System.Windows.Forms.Label label8;
        public System.Windows.Forms.TextBox textBox3;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.ToolStripMenuItem 报告类型ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 风险分析与SILToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem sIL分析报告ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem sIL验证报告ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 误停车报告ToolStripMenuItem;
    }
}

