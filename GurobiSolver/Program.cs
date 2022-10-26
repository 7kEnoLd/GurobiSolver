using HybridControl_FlexibleProducts.Information;
using HybridControl_FlexibleProducts.Models;
using NPOI.SS.Formula.Functions;
using NPOI.Util;
using System.Linq;
using MathNet.Numerics.Distributions;

namespace HybridControl_FlexibleProducts
{
    public class Program
    {
        static void Main(string[] args)
        {
            #region 导入Excel数据

            //导入OD维度的数据
            //包括OD的ID、名字、全价票价、需求均值、OD占用的区段资源
            string filepath = "C:\\Users\\oLd7kEn\\Desktop\\GurobiSolverData.xlsx";
            int sheetNum1 = 0;
            List<OD_Dimension> dt_ODDimension = OD_Dimension.ListAll(filepath, sheetNum1);

            //嵌套list储存分割后的各OD的区间使用情况
            List<List<int>> OD_UseSector = new List<List<int>>();
            for (int i = 0; i < dt_ODDimension.Count; i++)
            {
                string[] temp = dt_ODDimension[i].OD_UseSector.Split(",");
                List<int> ints = new List<int>();
                for (int j = 0; j < temp.Length; j++)
                {
                    ints.Add(Convert.ToInt32(temp[j]));
                }
                OD_UseSector.Add(ints);
            }

            //导入列车维度的数据
            //包括列车的ID、车次名、能够提供的OD
            int sheetNum2 = 1;
            List<Train_Dimension> dt_TrainDimension = Train_Dimension.ListAll(filepath, sheetNum2);

            //导入车站维度的数据
            //包括车站ID、车站名、里程数据
            int sheetNum3 = 2;
            List<Station_Dimension> dt_StationDimension = Station_Dimension.ListAll(filepath, sheetNum3);

            //导入车站基本数据
            //包括车站数、区段数、OD对数、席位容量、列车数
            int sheetNum4 = 3;
            List<StationConditions> dt_StationConditions = StationConditions.ListAll(filepath, sheetNum4);

            #endregion


            #region 处理各车次的可用矩阵,平分需求
            List<List<double>> demandMatrix = new();
            List<List<int>> countMatrix = new();
            List<List<double>> priceMatrix = new();
            for (int i = 0; i < dt_StationConditions[0].OD_Pair; i++)
            {
                List<int> countVector = new();
                List<double> priceVector = new();
                int temp = 0;
                for (int j = 0; j < dt_StationConditions[0].TrainNum; j++)
                {
                    List<Train_Dimension> station_Dimensions = dt_TrainDimension.FindAll(m => m.Train_ID == j + 1);
                    if (station_Dimensions.Exists(m => m.Train_UseOD == i + 1))
                    {
                        priceVector.Add(dt_ODDimension[i].OD_FullTicketPrice);
                        countVector.Add(1);
                        temp++;
                    }
                    else
                    {
                        countVector.Add(0);
                        priceVector.Add(0);
                    }
                }
                List<double> demandVector = new();
                countMatrix.Add(countVector);
                priceMatrix.Add(priceVector);
                double demand = dt_ODDimension[0].OD_CertainDemand / temp;
                for (int j = 0; j < dt_StationConditions[0].TrainNum; j++)
                {
                    List<Train_Dimension> station_Dimensions = dt_TrainDimension.FindAll(m => m.Train_ID == j + 1);
                    if (station_Dimensions.Exists(m => m.Train_UseOD == i + 1))
                    {
                        demandVector.Add(demand);
                    }
                    else
                    {
                        demandVector.Add(0);
                    }    
                }
                demandMatrix.Add(demandVector);
            }

            #endregion


            #region CDM(Certain Demand Model)模型调用Gurobi求解预定限额矩阵

            //CDM模型
            System.ValueTuple<List<List<double>>, double> Result2 = CertainDemandModelWithNoODCons.Optimize_CDM(priceMatrix, dt_ODDimension, dt_TrainDimension, dt_StationConditions, OD_UseSector, demandMatrix, countMatrix);
            #endregion

        }

    }

}