using ExportExcelFromOneDir.Lib;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace ExportExcelFromOneDir
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnSelect_Click(object sender, EventArgs e)
        {
            var dialogResult = folderBrowserDialog.ShowDialog();
            if (dialogResult == DialogResult.OK)
            {
                textBoxSelectPath.Text = folderBrowserDialog.SelectedPath;
            }
        }

        public void showMsg(string currentFile, int progressBarVal)
        {
            if (lbl_currentFile.InvokeRequired)
            {
                lbl_currentFile.Invoke(new Action<string>((m) =>
                {
                    lbl_currentFile.Text = "当前文件：" + currentFile;
                }), "");
            }
            else
            {
                lbl_currentFile.Text = "当前文件：" + currentFile;
            }

            if (progressBar.InvokeRequired)
            {
                progressBar.Invoke(new Action<string>((m) =>
                {
                    progressBar.Value = progressBarVal;
                }), "");
            }
            else
            {
                progressBar.Value = progressBarVal;
            }

        }

        public void showResult(string msg)
        {
            if (txtMsg.InvokeRequired)
            {
                txtMsg.Invoke(new Action<string>((m) =>
                {
                    txtMsg.Text = msg;
                }), "");
            }
            else
            {
                txtMsg.Text = msg;
            }
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            txtMsg.Text = "";
            if (string.IsNullOrEmpty(textBoxSelectPath.Text))
            {
                MessageBox.Show("请选择文件夹");
                return;
            }
            new Thread(() =>
            {
                try
                {

                    if (btnExport.InvokeRequired)
                    {
                        btnExport.Invoke(new Action<string>((m) =>
                        {
                            btnExport.Enabled = false;
                            btnExport.Text = "导出中...";
                        }), "");
                    }
                    else
                    {
                        btnExport.Enabled = true;
                        btnExport.Text = "导出中...";
                    }


                    //DirHelper dirHelper = new DirHelper(textBoxSelectPath.Text);
                    //List<string> fileNames = dirHelper.GetFileNames(showMsg);
                    ////MessageBox.Show(string.Join(",", fileNames.ToArray()));

                    //new ExcelHelper("所有文件名", fileNames, 1).ExportEXCEL();

                    ReadExcelHelper readExcelHelper = new ReadExcelHelper();
                    readExcelHelper.Read(txt_excelInfo.Text);
                    ExcelData excelData = readExcelHelper.GetData();

                    //MessageBox.Show("Test");
                    string msg = new CopDirHelper().handle(textBoxSelectPath.Text, excelData.Datas, showMsg);
                    showResult(msg);

                    DirectoryInfo path = new DirectoryInfo(textBoxSelectPath.Text);
                    Process.Start(path.Parent.FullName);

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                finally
                {
                    if (btnExport.InvokeRequired)
                    {
                        btnExport.Invoke(new Action<string>((m) =>
                        {
                            btnExport.Enabled = true;
                            btnExport.Text = "开始导出";
                        }), "");
                    }
                    else
                    {
                        btnExport.Enabled = true;
                        btnExport.Text = "开始导出";
                    }
                }
            }).Start();
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start("https://support.qq.com/products/449283/");
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        // 选择 excel 文件
        private void button1_Click(object sender, EventArgs e)
        {
            var dialogResult = openFileDialog1.ShowDialog();
            if (dialogResult == DialogResult.OK)
            {
                txt_excelInfo.Text = openFileDialog1.FileName;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void progressBar_Click(object sender, EventArgs e)
        {

        }
    }
}
