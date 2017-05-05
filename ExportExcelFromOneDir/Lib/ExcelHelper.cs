using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;

namespace ExportExcelFromOneDir.Lib
{
    public class ExcelHelper
    {
        private List<string> _data;
        private string _name;
        private int _colNum = 5;

        public ExcelHelper(string name, List<string> data, int colNum)
        {
            _name = string.Format("{0}_{1}.xls", name, DateTime.Now.ToString("yyyy_MM_dd_HH_mm_ss"));
            _data = data;
            _colNum = colNum;
        }

        public void ExportEXCEL()
        {
            var workbook = new HSSFWorkbook();
            var table = workbook.CreateSheet();            

            IRow row = null;
            for (int i = 0; i < _data.Count; i++)
            {
                var item = _data[i];

                if (i % _colNum == 0)
                    row = table.CreateRow((int)(i / _colNum));
                var cell = row.CreateCell(i % _colNum);
                cell.SetCellValue(item);
            }

            AutoCellWidth(table, (int)(_data.Count / _colNum));

            var savePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), _name);

            using (var fs = File.Open(savePath, FileMode.OpenOrCreate))
            {
                workbook.Write(fs);                
            }

            Process.Start(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));
        }

        private void AutoCellWidth(ISheet paymentSheet,int rowsCount)
        {
            for (int columnNum = 0; columnNum <= rowsCount; columnNum++)
            {
                int columnWidth = paymentSheet.GetColumnWidth(columnNum) / 256;
                for (int rowNum = 1; rowNum <= paymentSheet.LastRowNum; rowNum++)
                {
                    IRow currentRow;
                    //当前行未被使用过
                    if (paymentSheet.GetRow(rowNum) == null)
                    {
                        currentRow = paymentSheet.CreateRow(rowNum);
                    }
                    else
                    {
                        currentRow = paymentSheet.GetRow(rowNum);
                    }

                    if (currentRow.GetCell(columnNum) != null)
                    {
                        ICell currentCell = currentRow.GetCell(columnNum);
                        int length = Encoding.Default.GetBytes(currentCell.ToString()).Length;
                        if (columnWidth < length)
                        {
                            columnWidth = length;
                        }
                    }
                }
                paymentSheet.SetColumnWidth(columnNum, columnWidth * 256);
            }
        }
    }
}
