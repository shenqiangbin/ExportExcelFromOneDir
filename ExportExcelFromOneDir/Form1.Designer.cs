namespace ExportExcelFromOneDir
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.folderBrowserDialog = new System.Windows.Forms.FolderBrowserDialog();
            this.btnSelect = new System.Windows.Forms.Button();
            this.textBoxSelectPath = new System.Windows.Forms.TextBox();
            this.btnExport = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.toolStripStatusLabel1 = new System.Windows.Forms.ToolStripStatusLabel();
            this.progressBar = new System.Windows.Forms.ProgressBar();
            this.lbl_currentFile = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.linkLabel1 = new System.Windows.Forms.LinkLabel();
            this.label4 = new System.Windows.Forms.Label();
            this.txt_excelInfo = new System.Windows.Forms.TextBox();
            this.button1 = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.txtMsg = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.statusStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnSelect
            // 
            this.btnSelect.Location = new System.Drawing.Point(413, 86);
            this.btnSelect.Name = "btnSelect";
            this.btnSelect.Size = new System.Drawing.Size(75, 23);
            this.btnSelect.TabIndex = 0;
            this.btnSelect.Text = "选择";
            this.btnSelect.UseVisualStyleBackColor = true;
            this.btnSelect.Click += new System.EventHandler(this.btnSelect_Click);
            // 
            // textBoxSelectPath
            // 
            this.textBoxSelectPath.Location = new System.Drawing.Point(84, 87);
            this.textBoxSelectPath.Name = "textBoxSelectPath";
            this.textBoxSelectPath.Size = new System.Drawing.Size(323, 21);
            this.textBoxSelectPath.TabIndex = 1;
            // 
            // btnExport
            // 
            this.btnExport.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnExport.Location = new System.Drawing.Point(23, 116);
            this.btnExport.Name = "btnExport";
            this.btnExport.Size = new System.Drawing.Size(465, 39);
            this.btnExport.TabIndex = 2;
            this.btnExport.Text = "开始";
            this.btnExport.UseVisualStyleBackColor = true;
            this.btnExport.Click += new System.EventHandler(this.btnExport_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(25, 90);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(41, 12);
            this.label1.TabIndex = 3;
            this.label1.Text = "文件夹";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(33, 13);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(95, 12);
            this.label2.TabIndex = 4;
            this.label2.Text = "说明 ：改名程序";
            // 
            // statusStrip1
            // 
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripStatusLabel1});
            this.statusStrip1.Location = new System.Drawing.Point(0, 498);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Size = new System.Drawing.Size(1114, 22);
            this.statusStrip1.TabIndex = 5;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // toolStripStatusLabel1
            // 
            this.toolStripStatusLabel1.Name = "toolStripStatusLabel1";
            this.toolStripStatusLabel1.Size = new System.Drawing.Size(351, 17);
            this.toolStripStatusLabel1.Text = "作者：sqber，如有问题，请联系 QQ：1969858717（Turbo）";
            // 
            // progressBar
            // 
            this.progressBar.Location = new System.Drawing.Point(626, 218);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(465, 21);
            this.progressBar.TabIndex = 6;
            this.progressBar.Visible = false;
            this.progressBar.Click += new System.EventHandler(this.progressBar_Click);
            // 
            // lbl_currentFile
            // 
            this.lbl_currentFile.AutoSize = true;
            this.lbl_currentFile.Location = new System.Drawing.Point(21, 177);
            this.lbl_currentFile.Name = "lbl_currentFile";
            this.lbl_currentFile.Size = new System.Drawing.Size(65, 12);
            this.lbl_currentFile.TabIndex = 7;
            this.lbl_currentFile.Text = "当前文件：";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(9, 475);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(65, 12);
            this.label3.TabIndex = 8;
            this.label3.Text = "反馈地址：";
            // 
            // linkLabel1
            // 
            this.linkLabel1.AutoSize = true;
            this.linkLabel1.Location = new System.Drawing.Point(70, 475);
            this.linkLabel1.Name = "linkLabel1";
            this.linkLabel1.Size = new System.Drawing.Size(239, 12);
            this.linkLabel1.TabIndex = 9;
            this.linkLabel1.TabStop = true;
            this.linkLabel1.Text = "https://support.qq.com/products/449283/";
            this.linkLabel1.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel1_LinkClicked);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(25, 55);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(53, 12);
            this.label4.TabIndex = 10;
            this.label4.Text = "信息表格";
            // 
            // txt_excelInfo
            // 
            this.txt_excelInfo.Location = new System.Drawing.Point(84, 52);
            this.txt_excelInfo.Name = "txt_excelInfo";
            this.txt_excelInfo.Size = new System.Drawing.Size(323, 21);
            this.txt_excelInfo.TabIndex = 11;
            this.txt_excelInfo.TextChanged += new System.EventHandler(this.textBox1_TextChanged);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(413, 50);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 12;
            this.button1.Text = "选择";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // txtMsg
            // 
            this.txtMsg.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.txtMsg.Location = new System.Drawing.Point(23, 245);
            this.txtMsg.Multiline = true;
            this.txtMsg.Name = "txtMsg";
            this.txtMsg.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.txtMsg.Size = new System.Drawing.Size(1064, 208);
            this.txtMsg.TabIndex = 13;
            this.txtMsg.WordWrap = false;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(21, 218);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(41, 12);
            this.label5.TabIndex = 14;
            this.label5.Text = "结果：";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1114, 520);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.txtMsg);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.txt_excelInfo);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.linkLabel1);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.lbl_currentFile);
            this.Controls.Add(this.progressBar);
            this.Controls.Add(this.statusStrip1);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnExport);
            this.Controls.Add(this.textBoxSelectPath);
            this.Controls.Add(this.btnSelect);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "根据信息改名";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog;
        private System.Windows.Forms.Button btnSelect;
        private System.Windows.Forms.TextBox textBoxSelectPath;
        private System.Windows.Forms.Button btnExport;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel1;
        private System.Windows.Forms.ProgressBar progressBar;
        private System.Windows.Forms.Label lbl_currentFile;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.LinkLabel linkLabel1;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txt_excelInfo;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.TextBox txtMsg;
        private System.Windows.Forms.Label label5;
    }
}

