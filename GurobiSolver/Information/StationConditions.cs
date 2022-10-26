using ExcelWriteData;
using ExcelWriter_Utility;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Numerics;
using System.Text;
using System.Threading.Tasks;

namespace HybridControl_FlexibleProducts.Information
{
    public class StationConditions
    {
        public int Station { get; set; }

        public int? Sector
        {
            get
            {
                return Station - 1;
            }
            set
            {
                //区段数也只与车站数绑定
                value = Station - 1;
            }

        }


        private int _OD_Pair = 0;
        public int OD_Pair
        {
            get
            {
                if (Station is 0 || Station is 1)
                {
                    throw new Exception("请检查输入的车站数");
                }
                return _OD_Pair;
            }
            set
            {
                //确保这个OD值无法被修改，只能与车站数绑定
                _OD_Pair = value;
            }
        }

        //用构造器限制OD数，会报错因为OD_Pair类型无法识别
        public StationConditions() => _OD_Pair = CombinatorialNumber(Station);

        #region 车站数计算OD数
        /// <summary>
        /// 计算组合数，目的是为了根据车站计算OD数
        /// </summary>
        /// <param name="station">车站数</param>
        /// <param name="chooseNumber">在运输问题中，这个值取2</param>
        /// <returns>返回组合数，这个值等于OD数</returns>
        private int CombinatorialNumber(int station, int chooseNumber = 2)
        {
            if (chooseNumber is 0)
                return 1;
            int temp = (Factorial(station)) / (Factorial(chooseNumber) * Factorial(station - chooseNumber));
            return temp;
        }

        /// <summary>
        /// 计算阶乘数
        /// </summary>
        /// <param name="n">阶乘底数，这里用的int，如果数据过大可以换成BigInteger</param>
        /// <returns>返回阶乘数</returns>
        private int Factorial(int n)
        {
            if (n is 0)
                return 1;
            int temp = 1;
            for (int i = 1; i <= n; i++)
            {
                temp = temp * i;
            }
            return temp;
        }
        #endregion

        public int Capacity { get; set; }
        public int TrainNum { get; set; }


        /// <summary>
        /// 得到excel数据，并转化为与属性格式一致的list集合
        /// </summary>
        /// <param name="FileName"></param>
        /// <returns></returns>
        public static List<StationConditions> ListAll(string FileName, int sheetNum)
        {
            List<StationConditions> stationConditions = new List<StationConditions>();
            DataTable dt = LoadExcelData.ExcelToTable(FileName, sheetNum);
            foreach (DataRow dr in dt.Rows)
            {
                //DataRowToModel是扩展方法，参数用this修饰，因而在使用参数时可以直接参数.方法名格式
                stationConditions.Add(dr.DataRowToModel<StationConditions>());
            }
            return stationConditions;
        }
    }
}
