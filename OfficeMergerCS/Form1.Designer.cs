namespace OfficeMergerCS
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
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.btExport = new System.Windows.Forms.Button();
            this.tbDataTag = new System.Windows.Forms.TextBox();
            this.fileSystemWatcher1 = new System.IO.FileSystemWatcher();
            this.btnOpen = new System.Windows.Forms.Button();
            this.lvFile = new System.Windows.Forms.ListView();
            this.imageList1 = new System.Windows.Forms.ImageList(this.components);
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.youFindeToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.tbAdd = new System.Windows.Forms.TextBox();
            this.btnAdd = new System.Windows.Forms.Button();
            this.lbContent = new System.Windows.Forms.ListBox();
            this.button3 = new System.Windows.Forms.Button();
            this.btRead = new System.Windows.Forms.Button();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.tbMainEnd = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.tbMainStart = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.tbMainRange = new System.Windows.Forms.TextBox();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.label5 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.tbSheetCont = new System.Windows.Forms.TextBox();
            this.tbSheetPos = new System.Windows.Forms.TextBox();
            this.cbSheetSelect = new System.Windows.Forms.ComboBox();
            this.btClear = new System.Windows.Forms.Button();
            this.btReadWord = new System.Windows.Forms.Button();
            this.btReplace = new System.Windows.Forms.Button();
            this.btAutoEdit = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.fileSystemWatcher1)).BeginInit();
            this.contextMenuStrip1.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.SuspendLayout();
            // 
            // btExport
            // 
            this.btExport.Location = new System.Drawing.Point(391, 376);
            this.btExport.Name = "btExport";
            this.btExport.Size = new System.Drawing.Size(75, 23);
            this.btExport.TabIndex = 9;
            this.btExport.Text = "生成表格";
            this.btExport.UseVisualStyleBackColor = true;
            this.btExport.Click += new System.EventHandler(this.button1_Click);
            // 
            // tbDataTag
            // 
            this.tbDataTag.Location = new System.Drawing.Point(17, 175);
            this.tbDataTag.Name = "tbDataTag";
            this.tbDataTag.Size = new System.Drawing.Size(119, 21);
            this.tbDataTag.TabIndex = 6;
            this.tbDataTag.TextChanged += new System.EventHandler(this.textBox1_TextChanged);
            // 
            // fileSystemWatcher1
            // 
            this.fileSystemWatcher1.EnableRaisingEvents = true;
            this.fileSystemWatcher1.SynchronizingObject = this;
            // 
            // btnOpen
            // 
            this.btnOpen.Location = new System.Drawing.Point(21, 12);
            this.btnOpen.Name = "btnOpen";
            this.btnOpen.Size = new System.Drawing.Size(66, 48);
            this.btnOpen.TabIndex = 2;
            this.btnOpen.Text = "OPEN";
            this.btnOpen.UseVisualStyleBackColor = true;
            this.btnOpen.Click += new System.EventHandler(this.btnOpen_Click);
            // 
            // lvFile
            // 
            this.lvFile.Location = new System.Drawing.Point(21, 63);
            this.lvFile.Name = "lvFile";
            this.lvFile.Size = new System.Drawing.Size(261, 307);
            this.lvFile.TabIndex = 3;
            this.lvFile.UseCompatibleStateImageBehavior = false;
            this.lvFile.MouseClick += new System.Windows.Forms.MouseEventHandler(this.listView1_MouseClick);
            // 
            // imageList1
            // 
            this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
            this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
            this.imageList1.Images.SetKeyName(0, "Excel.jpg");
            this.imageList1.Images.SetKeyName(1, "Word.jpg");
            // 
            // contextMenuStrip1
            // 
            this.contextMenuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.youFindeToolStripMenuItem});
            this.contextMenuStrip1.Name = "contextMenuStrip1";
            this.contextMenuStrip1.Size = new System.Drawing.Size(150, 26);
            // 
            // youFindeToolStripMenuItem
            // 
            this.youFindeToolStripMenuItem.Name = "youFindeToolStripMenuItem";
            this.youFindeToolStripMenuItem.Size = new System.Drawing.Size(149, 22);
            this.youFindeToolStripMenuItem.Text = "You Find Me";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.tbAdd);
            this.groupBox1.Controls.Add(this.btnAdd);
            this.groupBox1.Controls.Add(this.lbContent);
            this.groupBox1.Location = new System.Drawing.Point(454, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(156, 357);
            this.groupBox1.TabIndex = 4;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "多点读取";
            // 
            // tbAdd
            // 
            this.tbAdd.Location = new System.Drawing.Point(30, 265);
            this.tbAdd.Name = "tbAdd";
            this.tbAdd.Size = new System.Drawing.Size(105, 21);
            this.tbAdd.TabIndex = 2;
            this.tbAdd.TextChanged += new System.EventHandler(this.tbAdd_TextChanged);
            this.tbAdd.KeyDown += new System.Windows.Forms.KeyEventHandler(this.tbAdd_KeyDown);
            // 
            // btnAdd
            // 
            this.btnAdd.Location = new System.Drawing.Point(30, 291);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(105, 33);
            this.btnAdd.TabIndex = 1;
            this.btnAdd.Text = "添加读取范围";
            this.btnAdd.UseVisualStyleBackColor = true;
            this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click);
            // 
            // lbContent
            // 
            this.lbContent.FormattingEnabled = true;
            this.lbContent.ItemHeight = 12;
            this.lbContent.Location = new System.Drawing.Point(30, 20);
            this.lbContent.Name = "lbContent";
            this.lbContent.Size = new System.Drawing.Size(105, 220);
            this.lbContent.TabIndex = 0;
            this.lbContent.SelectedIndexChanged += new System.EventHandler(this.lbContent_SelectedIndexChanged);
            this.lbContent.DoubleClick += new System.EventHandler(this.listBox1_DoubleClick);
            this.lbContent.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.listBox1_KeyPress);
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(38, 376);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(91, 52);
            this.button3.TabIndex = 3;
            this.button3.Text = "保存配置";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // btRead
            // 
            this.btRead.Location = new System.Drawing.Point(290, 376);
            this.btRead.Name = "btRead";
            this.btRead.Size = new System.Drawing.Size(95, 23);
            this.btRead.TabIndex = 8;
            this.btRead.Text = "读取内容";
            this.btRead.UseVisualStyleBackColor = true;
            this.btRead.Click += new System.EventHandler(this.btRead_Click);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.label3);
            this.groupBox2.Controls.Add(this.label6);
            this.groupBox2.Controls.Add(this.tbMainEnd);
            this.groupBox2.Controls.Add(this.label2);
            this.groupBox2.Controls.Add(this.tbMainStart);
            this.groupBox2.Controls.Add(this.label1);
            this.groupBox2.Controls.Add(this.tbMainRange);
            this.groupBox2.Controls.Add(this.tbDataTag);
            this.groupBox2.Location = new System.Drawing.Point(292, 148);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(156, 221);
            this.groupBox2.TabIndex = 5;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "区块读取";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(34, 110);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(89, 12);
            this.label3.TabIndex = 7;
            this.label3.Text = "结束单元格内容";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(52, 158);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(53, 12);
            this.label6.TabIndex = 8;
            this.label6.Text = "数据备注";
            // 
            // tbMainEnd
            // 
            this.tbMainEnd.Location = new System.Drawing.Point(17, 129);
            this.tbMainEnd.Name = "tbMainEnd";
            this.tbMainEnd.Size = new System.Drawing.Size(119, 21);
            this.tbMainEnd.TabIndex = 5;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(34, 65);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(89, 12);
            this.label2.TabIndex = 5;
            this.label2.Text = "标题单元格内容";
            // 
            // tbMainStart
            // 
            this.tbMainStart.Location = new System.Drawing.Point(17, 83);
            this.tbMainStart.Name = "tbMainStart";
            this.tbMainStart.Size = new System.Drawing.Size(119, 21);
            this.tbMainStart.TabIndex = 4;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(40, 19);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(77, 12);
            this.label1.TabIndex = 3;
            this.label1.Text = "重点读取区域";
            // 
            // tbMainRange
            // 
            this.tbMainRange.Location = new System.Drawing.Point(17, 37);
            this.tbMainRange.Name = "tbMainRange";
            this.tbMainRange.Size = new System.Drawing.Size(119, 21);
            this.tbMainRange.TabIndex = 2;
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.label5);
            this.groupBox3.Controls.Add(this.label4);
            this.groupBox3.Controls.Add(this.tbSheetCont);
            this.groupBox3.Controls.Add(this.tbSheetPos);
            this.groupBox3.Controls.Add(this.cbSheetSelect);
            this.groupBox3.Location = new System.Drawing.Point(292, 12);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(156, 119);
            this.groupBox3.TabIndex = 6;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "活页簿";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(85, 64);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(53, 12);
            this.label5.TabIndex = 4;
            this.label5.Text = "对应内容";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(15, 64);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(65, 12);
            this.label4.TabIndex = 3;
            this.label4.Text = "特征单元格";
            // 
            // tbSheetCont
            // 
            this.tbSheetCont.Location = new System.Drawing.Point(87, 82);
            this.tbSheetCont.Name = "tbSheetCont";
            this.tbSheetCont.Size = new System.Drawing.Size(51, 21);
            this.tbSheetCont.TabIndex = 2;
            // 
            // tbSheetPos
            // 
            this.tbSheetPos.Location = new System.Drawing.Point(17, 82);
            this.tbSheetPos.Name = "tbSheetPos";
            this.tbSheetPos.Size = new System.Drawing.Size(49, 21);
            this.tbSheetPos.TabIndex = 1;
            // 
            // cbSheetSelect
            // 
            this.cbSheetSelect.FormattingEnabled = true;
            this.cbSheetSelect.Items.AddRange(new object[] {
            "全部",
            "1",
            "2",
            "3",
            "4",
            "5",
            "6"});
            this.cbSheetSelect.Location = new System.Drawing.Point(15, 28);
            this.cbSheetSelect.Name = "cbSheetSelect";
            this.cbSheetSelect.Size = new System.Drawing.Size(121, 20);
            this.cbSheetSelect.TabIndex = 0;
            this.cbSheetSelect.Text = "全部";
            // 
            // btClear
            // 
            this.btClear.Location = new System.Drawing.Point(143, 12);
            this.btClear.Name = "btClear";
            this.btClear.Size = new System.Drawing.Size(66, 48);
            this.btClear.TabIndex = 7;
            this.btClear.Text = "Clear";
            this.btClear.UseVisualStyleBackColor = true;
            this.btClear.Click += new System.EventHandler(this.btClear_Click);
            // 
            // btReadWord
            // 
            this.btReadWord.Location = new System.Drawing.Point(290, 405);
            this.btReadWord.Name = "btReadWord";
            this.btReadWord.Size = new System.Drawing.Size(95, 23);
            this.btReadWord.TabIndex = 8;
            this.btReadWord.Text = "读取Word";
            this.btReadWord.UseVisualStyleBackColor = true;
            this.btReadWord.Click += new System.EventHandler(this.btReadWord_Click);
            // 
            // btReplace
            // 
            this.btReplace.Location = new System.Drawing.Point(391, 406);
            this.btReplace.Name = "btReplace";
            this.btReplace.Size = new System.Drawing.Size(75, 23);
            this.btReplace.TabIndex = 9;
            this.btReplace.Text = "替换文字";
            this.btReplace.UseVisualStyleBackColor = true;
            this.btReplace.Click += new System.EventHandler(this.btReplace_Click);
            // 
            // btAutoEdit
            // 
            this.btAutoEdit.Location = new System.Drawing.Point(477, 405);
            this.btAutoEdit.Name = "btAutoEdit";
            this.btAutoEdit.Size = new System.Drawing.Size(95, 23);
            this.btAutoEdit.TabIndex = 10;
            this.btAutoEdit.Text = "批量修改";
            this.btAutoEdit.UseVisualStyleBackColor = true;
            this.btAutoEdit.Click += new System.EventHandler(this.btAutoEdit_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(618, 443);
            this.Controls.Add(this.btAutoEdit);
            this.Controls.Add(this.btReplace);
            this.Controls.Add(this.btReadWord);
            this.Controls.Add(this.btClear);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.btRead);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.lvFile);
            this.Controls.Add(this.btnOpen);
            this.Controls.Add(this.btExport);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form1";
            this.Text = "表格合并器";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.fileSystemWatcher1)).EndInit();
            this.contextMenuStrip1.ResumeLayout(false);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btExport;
        private System.Windows.Forms.TextBox tbDataTag;
        private System.IO.FileSystemWatcher fileSystemWatcher1;
        private System.Windows.Forms.Button btnOpen;
        private System.Windows.Forms.ListView lvFile;
        private System.Windows.Forms.ImageList imageList1;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
        private System.Windows.Forms.ToolStripMenuItem youFindeToolStripMenuItem;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.TextBox tbAdd;
        private System.Windows.Forms.Button btnAdd;
        private System.Windows.Forms.ListBox lbContent;
        private System.Windows.Forms.Button btRead;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox tbMainEnd;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox tbMainStart;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox tbMainRange;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.ComboBox cbSheetSelect;
        private System.Windows.Forms.TextBox tbSheetCont;
        private System.Windows.Forms.TextBox tbSheetPos;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button btClear;
        private System.Windows.Forms.Button btReadWord;
        private System.Windows.Forms.Button btReplace;
        private System.Windows.Forms.Button btAutoEdit;
        private System.Windows.Forms.Label label6;
    }
}

