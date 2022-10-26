using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;

namespace ExcelWriteData
{
    public class ExportExcelData
    {
        /// <summary>
        /// 输出excel格式文件，但是即使是xlsx格式，具体单元格格式用xls的也可以用，先暂时用着，等报错了再说
        /// </summary>
        /// <param name="table"></param>
        /// <param name="fileName"></param>
        /// <param name="workbook"></param>
        /// <param name="sheetName"></param>
        public static void ExportExcel(DataTable table, string fileName, IWorkbook workbook, string sheetName)
        {
            //设置表名
            ISheet sheet = workbook.CreateSheet(sheetName);
            ICellStyle headercellStyle = GetHeaderStyle(workbook);

            NPOI.SS.UserModel.IFont cellfont = workbook.CreateFont();
            cellfont.IsBold = false;
            cellfont.FontName = "宋体";
            cellfont.FontHeightInPoints = 11;

            //修改插入excel的各列数据格式
            ICellStyle cellStyle = GetCellStyle(workbook);
            cellStyle.DataFormat = HSSFDataFormat.GetBuiltinFormat("@");
            cellStyle.SetFont(cellfont);

            ICellStyle numCellStyle = GetCellStyle(workbook);
            numCellStyle.SetFont(cellfont);
            numCellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Right;
            numCellStyle.DataFormat = HSSFDataFormat.GetBuiltinFormat("#,##0.00");

            ICellStyle ratioCellStyle = GetCellStyle(workbook);
            ratioCellStyle.SetFont(cellfont);
            ratioCellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Right;
            ratioCellStyle.DataFormat = HSSFDataFormat.GetBuiltinFormat("0.00%");

            int iRowIndex = 0;
            int icolIndex = 0;
            IRow headerRow = sheet.CreateRow(iRowIndex);
            foreach (DataColumn item in table.Columns)
            {
                ICell cell = headerRow.CreateCell(icolIndex);
                cell.SetCellValue(item.ColumnName);
                cell.CellStyle = headercellStyle;
                icolIndex++;
            }
            iRowIndex++;

            int iCellIndex = 0;
            foreach (DataRow row in table.Rows)
            {
                IRow DataRow = sheet.CreateRow(iRowIndex);
                foreach (DataColumn colItem in table.Columns)
                {
                    Type type = row[colItem].GetType();
                    //这里默认double和int类型
                    ICell cell = DataRow.CreateCell(iCellIndex);
                    if (type == typeof(double) || type == typeof(int))
                    {
                        cell.SetCellValue(ToDoubleEx(row[colItem]));
                        cell.CellStyle = numCellStyle;
                    }
                    //设置百分比格式
                    //else if (colItem.ColumnName.Contains("占比"))
                    //{
                    //    cell.SetCellValue(Convert.ToDouble(row[colItem]));
                    //    cell.CellStyle = ratioCellStyle;
                    //}
                    else
                    {
                        cell.SetCellValue(row[colItem].ToString());
                        cell.CellStyle = cellStyle;
                    }
                    iCellIndex++;
                }
                iCellIndex = 0;
                iRowIndex++;
            }

            List<int> colsLength = new List<int>();
            foreach (DataColumn column in table.Columns)
            {
                var length = table.AsEnumerable().Max(row => row[column].ToString().Length);
                colsLength.Add(length);
            }

            AutoColumnWidth(sheet, table.Columns.Count, colsLength.ToArray(), 9);
        }

        /// <summary>
        /// 设置列宽
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="cols"></param>
        /// <param name="colLength"></param>
        /// <param name="addlength"></param>
        private static void AutoColumnWidth(ISheet sheet, int cols, int[] colLength, int addlength)
        {
            for (int col = 0; col < cols; col++)
            {
                var columnWidth = colLength[col] * 256 + 30 * 256;
                sheet.SetColumnWidth(col, columnWidth);
            }
        }

        private static ICellStyle GetCellStyle(IWorkbook workbook)
        {
            ICellStyle cellStyle = workbook.CreateCellStyle();
            cellStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            cellStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            cellStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            cellStyle.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            return cellStyle;
        }

        private static ICellStyle GetHeaderStyle(IWorkbook workbook)
        {
            ICellStyle headercellStyle = workbook.CreateCellStyle();
            headercellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            headercellStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            headercellStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            headercellStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            headercellStyle.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;

            headercellStyle.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Grey25Percent.Index;
            headercellStyle.FillPattern = FillPattern.SolidForeground;
            NPOI.SS.UserModel.IFont headerfont = workbook.CreateFont();
            headerfont.IsBold = true;
            headerfont.FontName = "宋体";
            headerfont.FontHeightInPoints = 11;
            headercellStyle.SetFont(headerfont);

            return headercellStyle;
        }

        private static double ToDoubleEx(object obj)
        {
            if (obj == DBNull.Value)
            {
                return 0;
            }
            string str = obj.ToString();
            if (str == null || str.Trim() == string.Empty)
            {
                return 0;
            }
            else
            {
                return Convert.ToDouble(str);
            }
        }
    }
}
