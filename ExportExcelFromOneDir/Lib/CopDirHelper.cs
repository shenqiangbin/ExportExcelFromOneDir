using ExportExcelFromOneDir.model;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace ExportExcelFromOneDir.Lib
{
    public class CopDirHelper
    {
        private List<string> hasHandled = new List<string>();
        private StringBuilder builder = new StringBuilder();
        private string firstpath = "";
        private string firstNewPath = "";
        private Process_EventHandler processHandler;

        public delegate void Process_EventHandler(string currentFile, int processVal);


        public string handle(string path, List<List<object>> datas, Process_EventHandler processHandler)
        {
            builder = new StringBuilder();
            firstpath = path;
            firstNewPath = firstpath + "副本-程序创建-" + DateTime.Now.ToString("yyyy_MM_dd_HH_mm_ss");
            this.processHandler = processHandler;

            DirectoryInfo dir = new DirectoryInfo(path);
            //CopyDirectory(path, newpath);

            handleDir(path, datas);

            return builder.ToString();

        }

        private void handleDir(string path, List<List<object>> datas)
        {
            DirectoryInfo di = Directory.CreateDirectory(path);
            string[] files = Directory.GetFileSystemEntries(path);

            foreach (string file in files)
            {
                if (Directory.Exists(file))
                {                    
                    handleDir(file,datas);
                }
                else
                {
                    this.processHandler(file, 0);
                    FileInfo fileInfo = new FileInfo(file);
                    NewFile newFileInfo = ContainsFile(datas, di.Name, fileInfo.Name);

                    if (newFileInfo != null)
                    {
                        string floderName = fileInfo.DirectoryName;
                        floderName = floderName.Replace(firstpath, firstNewPath);
                        floderName = floderName.Replace(di.Name, newFileInfo.NewPath);

                        if (!Directory.Exists(floderName))
                        {
                            Directory.CreateDirectory(floderName);

                        }

                        string destFileName = Path.Combine(floderName, newFileInfo.NewName);

                        File.Copy(file, destFileName, true);

                        builder.AppendFormat("【{0}】处理后为【{1}】\r\n", file, destFileName);

                    }
                    else
                    {
                        builder.AppendFormat("【{0}】未处理\r\n", file);
                    }

                    //DirectoryInfo theDir = Directory.CreateDirectory(file);
                    //if (!hasHandled.Contains(theDir.Name))
                    //{
                    //    string newDirName = ContainsDir(datas, theDir.Name);
                    //    if (newDirName != null)
                    //    {
                    //        handleDir(theDir, datas);

                    //        //string floderName = theDir.Parent.FullName;
                    //        //string destName = Path.Combine(floderName, newDirName);
                    //        //Directory.Move(theDir.FullName, destName);

                    //        //hasHandled.Add(theDir.Name);

                    //    }
                    //}
                }
            }
        }

        private void handleDir(DirectoryInfo theDir, List<List<object>> datas)
        {
            //FileInfo[] fileInfos = theDir.GetFiles();
            //foreach (FileInfo file in fileInfos)
            //{
            //    string newFileName = ContainsFile(datas, theDir.Name, file.Name);
            //    if (newFileName != null)
            //    {
            //        string floderName = file.DirectoryName;
            //        string destFileName = Path.Combine(floderName, newFileName);
            //        File.Move(file.FullName, destFileName);
            //    }
            //    else
            //    {
            //        builder.AppendFormat("{0}未处理\r\n", file.FullName);
            //    }
            //}
        }

        private string ContainsDir(List<List<object>> datas, string dirName)
        {
            foreach (List<object> rows in datas)
            {
                if (rows[2].ToString() == dirName)
                {
                    return rows[0].ToString();
                }
            }
            return null;
        }

        private NewFile ContainsFile(List<List<object>> datas, string dirName, string fileName)
        {
            foreach (List<object> rows in datas)
            {
                if (rows[2].ToString() == dirName)
                {
                    string oldfile = rows[3].ToString().Trim();
                    if (fileName.Contains(oldfile))
                    {
                        string newFileName = fileName.Replace(oldfile, rows[1].ToString().Trim());
                        string newPath = rows[0].ToString().Trim();

                        NewFile newfile = new NewFile(newFileName, newPath);
                        return newfile;
                    }
                }
            }
            return null;
        }

        private void CopyDirectory(string sourcePath, string destPath)
        {
            string floderName = Path.GetFileName(sourcePath);
            DirectoryInfo di = Directory.CreateDirectory(Path.Combine(destPath, floderName));
            string[] files = Directory.GetFileSystemEntries(sourcePath);

            foreach (string file in files)
            {
                if (Directory.Exists(file))
                {
                    CopyDirectory(file, di.FullName);
                }
                else
                {
                    File.Copy(file, Path.Combine(di.FullName, Path.GetFileName(file)), true);
                }
            }
        }
    }
}
