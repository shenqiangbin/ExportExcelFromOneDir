using ExportExcelFromOneDir.Lib;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
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
                    lbl_currentFile.Text = currentFile;
                }), "");
            }
            else
            {
                lbl_currentFile.Text = currentFile;
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

        private void btnExport_Click(object sender, EventArgs e)
        {
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


                    DirHelper dirHelper = new DirHelper(textBoxSelectPath.Text);
                    List<string> fileNames = dirHelper.GetFileNames(showMsg);
                    //MessageBox.Show(string.Join(",", fileNames.ToArray()));

                    new ExcelHelper("所有文件名", fileNames, 1).ExportEXCEL();
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
    }
}
