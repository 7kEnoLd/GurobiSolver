using ExcelWriteData;
using ExcelWriter_Utility;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HybridControl_FlexibleProducts.Information
{
    public class Station_Dimension
    {
        public int Station_ID { get; set; }
        public string? Station_Name { get; set; }
        public int Station_Distance { get; set; }

        /// <summary>
        /// 得到excel数据，并转化为与属性格式一致的list集合
        /// </summary>
        /// <param name="FileName"></param>
        /// <returns></returns>
        public static List<Station_Dimension> ListAll(string FileName, int sheetNum)
        {
            List<Station_Dimension> station_Dimensions = new List<Station_Dimension>();
            DataTable dt = LoadExcelData.ExcelToTable(FileName, sheetNum);
            foreach (DataRow dr in dt.Rows)
            {
                //DataRowToModel是扩展方法，参数用this修饰，因而在使用参数时可以直接参数.方法名格式
                station_Dimensions.Add(dr.DataRowToModel<Station_Dimension>());
            }
            return station_Dimensions;
        }

    }
}
