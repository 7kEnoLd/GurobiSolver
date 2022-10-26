using Microsoft.VisualBasic;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;

using NPOI.XSSF.UserModel;
using System.Collections.Generic;
using System.Data;

namespace ExcelWriteData
{
    public class LoadExcelData
    {

        /// <summary>
        /// xxc定制优化后
        /// Excel导入成Datable
        /// </summary>
        /// <param name="file">导入路径(包含文件名与扩展名)</param>
        /// <returns></returns>
        public static DataTable ExcelToTable(string file, int sheetNum)
        {
            try
            {
                DataTable dt = new DataTable();
                IWorkbook workbook;
                string fileExt = Path.GetExtension(file).ToLower();
                using (FileStream fs = new FileStream(file, FileMode.Open, FileAccess.Read))
                {
                    //XSSFWorkbook 适用XLSX格式，HSSFWorkbook 适用XLS格式
                    if (fileExt == ".xlsx")
                    { workbook = new XSSFWorkbook(fs); }
                    else if (fileExt == ".xls")
                    { workbook = new HSSFWorkbook(fs); }
                    else { workbook = null; }
                    if (workbook == null)
                    { return null; }
                    ISheet sheet = workbook.GetSheetAt(sheetNum); //修改导入的表格位置

                    //表头  
                    IRow header = sheet.GetRow(sheet.FirstRowNum);
                    List<int> columns = new List<int>();
                    for (int i = 0; i < header.LastCellNum; i++)
                    {
                        object obj = GetValueType(file, workbook, header.GetCell(i));
                        if (obj == null || obj.ToString() == string.Empty)
                        {
                            //xxc20220613遇空列，停止读取列
                            break;
                            //dt.Columns.Add(new DataColumn("Columns" + i.ToString()));
                        }
                        else
                            dt.Columns.Add(new DataColumn(obj.ToString())); //添加表头列名
                        columns.Add(i);
                    }
                    //数据  
                    for (int i = sheet.FirstRowNum + 1; i <= sheet.LastRowNum; i++) //这里可以选择行范围，默认所有行
                    {
                        DataRow dr = dt.NewRow();
                        bool hasValue = false;
                        foreach (int j in columns)
                        {
                            try
                            {
                                dr[j] = GetValueType(file, workbook, sheet.GetRow(i).GetCell(j));
                                if (dr[j] != null && dr[j].ToString() != string.Empty)
                                {
                                    hasValue = true;
                                }
                            }
                            catch (Exception ee)
                            {
                                continue;
                            }

                        }

                        if (hasValue)
                        {
                            dt.Rows.Add(dr);
                        }
                        else
                        {
                            //xxc20220613遇空行，停止读取行
                            break;
                        }
                    }
                }
                return dt;
            }
            catch (Exception e)
            {
                return null;
            }

        }





        /// <summary>
        /// 获取单元格类型
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        private static object GetValueType(string strFileName, IWorkbook workbook, ICell cell)
        {
            if (cell == null)
                return null;
            switch (cell.CellType)
            {
                case CellType.Blank: //BLANK:  
                    return null;
                case CellType.Boolean: //BOOLEAN:  
                    return cell.BooleanCellValue;
                case CellType.Numeric: //NUMERIC:  
                    return cell.NumericCellValue;
                case CellType.String: //STRING:  
                    return cell.StringCellValue;
                case CellType.Error: //ERROR:  
                    return cell.ErrorCellValue;
                case CellType.Formula: //FORMULA:
                    object rv = null;
                    if (Path.GetExtension(strFileName).ToLower().Trim() == ".xlsx")
                    {
                        XSSFFormulaEvaluator eva = new XSSFFormulaEvaluator(workbook);
                        if (eva.Evaluate(cell).CellType == CellType.Numeric)
                        {
                            rv = eva.Evaluate(cell).NumberValue;
                        }
                        else
                        {
                            rv = eva.Evaluate(cell).StringValue;
                        }
                    }
                    else
                    {
                        HSSFFormulaEvaluator eva = new HSSFFormulaEvaluator(workbook);
                        if (eva.Evaluate(cell).CellType == CellType.Numeric)
                        {
                            rv = eva.Evaluate(cell).NumberValue;
                        }
                        else
                        {
                            rv = eva.Evaluate(cell).StringValue;
                        }
                    }
                    return rv;
                default:
                    return "=" + cell.CellFormula;
            }
        }
    }
}

