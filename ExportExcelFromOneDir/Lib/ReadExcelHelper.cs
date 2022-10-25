using NPOI.HSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;

namespace ExportExcelFromOneDir.Lib
{
    public class ReadExcelHelper
    {
        private IWorkbook workbook;

        public void Read(string filePath)
        {
            string fileExt = Path.GetExtension(filePath).ToLower();

            FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read);
            if (fileExt == ".xlsx")
            {
                workbook = new XSSFWorkbook(fs);
            }
            else if (fileExt == ".xls")
            {
                workbook = new HSSFWorkbook(fs);
            }

        }

        public ExcelData GetData()
        {
            ExcelData excelData = new ExcelData();

            ISheet sheet = workbook.GetSheet("编号信息");
            if (sheet == null)
                throw new Exception("没有编号信息sheet页");

            //表头  
            IRow header = sheet.GetRow(sheet.FirstRowNum);
            List<int> columns = new List<int>();
            for (int i = 0; i < header.LastCellNum; i++)
            {
                object obj = GetValueType(header.GetCell(i));    
                excelData.Titles.Add(obj.ToString());
                columns.Add(i);
            }
            //数据  
            for (int i = sheet.FirstRowNum + 1; i <= sheet.LastRowNum; i++)
            {
                List<object> rowData = new List<object>();
                foreach (int j in columns)
                {
                    object obj = GetValueType(sheet.GetRow(i).GetCell(j));
                    if (obj != null && obj.ToString() != string.Empty)
                    {
                        rowData.Add(obj);
                    }
                    else
                    {
                        rowData.Add("");
                    }
                }

                excelData.Datas.Add(rowData);
              
            }

            return excelData;
        }

        /// <summary>
        /// 获取单元格类型
        /// </summary>
        /// <param name="cell">目标单元格</param>
        /// <returns></returns>
        private static object GetValueType(ICell cell)
        {
            if (cell == null)
                return null;
            switch (cell.CellType)
            {
                case CellType.Blank:
                    return null;
                case CellType.Boolean:
                    return cell.BooleanCellValue;
                case CellType.Numeric:
                    return cell.NumericCellValue;
                case CellType.String:
                    return cell.StringCellValue;
                case CellType.Error:
                    return cell.ErrorCellValue;
                case CellType.Formula:
                    return cell.StringCellValue;
                default:
                    return "=" + cell.CellFormula;
            }
        }
    }
}
