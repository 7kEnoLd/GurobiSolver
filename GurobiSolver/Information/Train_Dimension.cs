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
    public class Train_Dimension
    {
        public int Train_ID { get; set; }
        public string? Train_Name { get; set; }
        public int Train_UseOD { get; set; }

        /// <summary>
        /// 得到excel数据，并转化为与属性格式一致的list集合
        /// </summary>
        /// <param name="FileName"></param>
        /// <returns></returns>
        public static List<Train_Dimension> ListAll(string FileName, int sheetNum)
        {
            List<Train_Dimension> train_Dimensions = new List<Train_Dimension>();
            DataTable dt = LoadExcelData.ExcelToTable(FileName, sheetNum);
            foreach (DataRow dr in dt.Rows)
            {
                //DataRowToModel是扩展方法，参数用this修饰，因而在使用参数时可以直接参数.方法名格式
                train_Dimensions.Add(dr.DataRowToModel<Train_Dimension>());
            }
            return train_Dimensions;
        }
    }
}
