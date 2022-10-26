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
    public class OD_Dimension
    {
        public int OD_ID { get; set; }
        public string? OD_Name { get; set; }
        public double OD_FullTicketPrice { get; set; }
        public double OD_CertainDemand { get; set; }
        public string? OD_UseSector { get; set; }

        /// <summary>
        /// 得到excel数据，并转化为与属性格式一致的list集合
        /// </summary>
        /// <param name="FileName"></param>
        /// <returns></returns>
        public static List<OD_Dimension> ListAll(string FileName, int sheetNum)
        {
            List<OD_Dimension>  oD_Dimensions = new List<OD_Dimension>();
            DataTable dt = LoadExcelData.ExcelToTable(FileName, sheetNum);
            foreach (DataRow dr in dt.Rows)
            {
                //DataRowToModel是扩展方法，参数用this修饰，因而在使用参数时可以直接参数.方法名格式
                oD_Dimensions.Add(dr.DataRowToModel<OD_Dimension>());
            }
            return oD_Dimensions;
        }

    }
}
