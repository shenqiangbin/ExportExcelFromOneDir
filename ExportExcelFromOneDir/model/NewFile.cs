using System;
using System.Collections.Generic;
using System.Text;

namespace ExportExcelFromOneDir.model
{
    public class NewFile
    {
        public string NewName;
        public string NewPath;

        public NewFile(string newName, string newPath)
        {
            NewName = newName;
            NewPath = newPath;
        }
    }
}
