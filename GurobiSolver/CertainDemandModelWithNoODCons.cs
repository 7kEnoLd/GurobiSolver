using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Gurobi;
using HybridControl_FlexibleProducts.Information;
using NPOI.HSSF.Record;

namespace HybridControl_FlexibleProducts.Models
{
    public class CertainDemandModelWithNoODCons
    {

        public static System.ValueTuple<List<List<double>>, double> Optimize_CDM(List<List<double>> priceMatrix, List<OD_Dimension> odDimension, List<Train_Dimension> trainDimension, List<StationConditions> stationCondition, List<List<int>> odUseSector, List<List<double>> demandMatrix, List<List<int>> countMatrix)
        {
            try
            {
                //创建空环境，设置参数并开始
                GRBEnv env = new GRBEnv(true);
                env.Set("Logfile", "mip.log");
                env.Start();

                //创建一个空模型
                GRBModel model = new GRBModel(env);
                model.ModelName = "Optimize_CDM";

                //决策变量,目标函数
                GRBVar[,] x_BookingLimits = new GRBVar[stationCondition[0].OD_Pair, stationCondition[0].TrainNum];
                for (int i = 0; i < stationCondition[0].OD_Pair; i++)
                {
                    double lb = 0;
                    double ub = 10000;
                    for (int j = 0; j < stationCondition[0].TrainNum; j++)
                    {
                        string name1 = (i + 1).ToString() + "," + (j + 1).ToString();
                        x_BookingLimits[i, j] = model.AddVar(lb, ub, priceMatrix[i][j], GRB.INTEGER, name1);
                    }
                }

                //约束条件1：能力约束,这里放列车的容量
                //只要提供服务的约束
                for (int i = 0; i < stationCondition[0].Sector; i++)
                {
                    for (int j = 0; j < stationCondition[0].TrainNum; j++)
                    {
                        GRBLinExpr capacityRestrict = 0.0;
                        for (int k = 0; k < stationCondition[0].OD_Pair; k++)
                        {
                            if (odUseSector[k].Exists(m => m == i + 1) && countMatrix[k][j] is not 0)
                            {
                                GRBVar gRBVarX = x_BookingLimits[k, j];
                                capacityRestrict.AddTerm(1.0, gRBVarX);
                            }
                        }
                        model.AddConstr(capacityRestrict <= stationCondition[0].Capacity, "能力约束：" + (i + 1).ToString() + "，" + (j + 1).ToString());
                    }

                }

                //约束条件2：需求约束，确定需求
                //只要有需求的条件
                for (int i = 0; i < stationCondition[0].OD_Pair; i++)
                {
                    if (odDimension[i].OD_CertainDemand is not 0)
                    {
                        for (int j = 0; j < stationCondition[0].TrainNum; j++)
                        {
                            if (countMatrix[i][j] is not 0)
                            {
                                GRBLinExpr demandRestrictX = 0.0;
                                GRBVar gRBVarX = x_BookingLimits[i, j];
                                demandRestrictX.AddTerm(1.0, gRBVarX);
                                model.AddConstr(demandRestrictX <= demandMatrix[i][j], "确定需求约束：" + (i + 1).ToString() + "，" + (j + 1).ToString());
                            }
                        }
                    }
                }



                //设置目标类型，并优化
                model.ModelSense = GRB.MAXIMIZE;
                model.Optimize();

                //判断是否求出最优解，这里的判断条件不太清楚
                if (model.Status != GRB.Status.OPTIMAL)
                {
                    Console.WriteLine("尚未求出最优解");
                }

                List<List<double>> xResult = new();
                for (int i = 0; i < stationCondition[0].OD_Pair; i++)
                {
                    List<double> xResultOne = new();
                    for (int j = 0; j < stationCondition[0].TrainNum; j++)
                    {
                        xResultOne.Add(x_BookingLimits[i, j].X);
                    }
                    xResult.Add(xResultOne);

                }

                //返回结果
                return (xResult, model.ObjVal);

                //释放模型和环境的非托管资源
                model.Dispose();
                env.Dispose();
            }
            catch (GRBException e)
            {
                throw new Exception("Error code:" + e.ErrorCode + "." + e.Message);
            }
        }



    }
}
