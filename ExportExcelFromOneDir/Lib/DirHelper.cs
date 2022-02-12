using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace ExportExcelFromOneDir.Lib
{
    public class DirHelper
    {
        private string _dirPath;
        private List<string> _extensions = new List<string>();

        public delegate void Process_EventHandler(string currentFile, int processVal);

        public DirHelper(string dirPath)
        {
            _dirPath = dirPath;
            _extensions.AddRange(new string[] { ".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx", ".pdf" });
        }

        public List<string> GetFileNames(Process_EventHandler processHandler)
        {
            List<string> fileNames = new List<string>();

            //DirectoryInfo dir = new DirectoryInfo(_dirPath);
            //FileInfo[] files = dir.GetFiles("*.*", SearchOption.AllDirectories);
            List<FileInfo> files = GetFiles(_dirPath, "*.*");
            int i = 1;
            int total = files.Count;
            foreach (var file in files)
            {
                if (_extensions.Contains(file.Extension.ToLower()))
                {
                    fileNames.Add(file.FullName);
                }

                int processVal = (int)(((double)i / total) * 100);
                string currentFile = string.Format("({0}/{1}){2}", i, total, file.Name);
                processHandler(currentFile, processVal);
                i++;

            }

            return fileNames;
        }

        private List<FileInfo> GetFiles(string path, string pattern)
        {
            var files = new List<FileInfo>();

            try
            {
                DirectoryInfo dir = new DirectoryInfo(path);
                files.AddRange(dir.GetFiles(pattern, SearchOption.TopDirectoryOnly));
                foreach (var directory in Directory.GetDirectories(path))
                    files.AddRange(GetFiles(directory, pattern));
            }
            catch (UnauthorizedAccessException) { }

            return files;
        }

        private List<FileInfo> GetFiles(string path, List<String> patterns)
        {
            var files = new List<FileInfo>();

            try
            {
                DirectoryInfo dir = new DirectoryInfo(path);
                foreach (string pattern in patterns)
                {
                    files.AddRange(dir.GetFiles(pattern, SearchOption.TopDirectoryOnly));
                }
                foreach (var directory in Directory.GetDirectories(path))
                    files.AddRange(GetFiles(directory, patterns));
            }
            catch (UnauthorizedAccessException) { }

            return files;
        }

        internal List<string> GetFileNames(object showMsg)
        {
            throw new NotImplementedException();
        }
    }
}
