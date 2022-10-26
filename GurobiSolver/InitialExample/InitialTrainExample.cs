using HybridControl_FlexibleProducts.Information;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HybridControl_FlexibleProducts.InitialExample
{
    public class InitialTrainExample
    {
        public int[,,] capacityMatrix;

        /// <summary>
        /// 初始化所有列车席位资源
        /// </summary>
        /// <param name="sample">样本数量</param>
        /// <param name="capacity">席位数量</param>
        /// <param name="sector">区段</param>
        /// <param name="trainNum">车次</param>
        public InitialTrainExample(int sample, int capacity, int sector, int trainNum)
        {
            //初始化所有车次所有区间所有席位的资源可用性，置为1
            capacityMatrix = new int[capacity, sector, trainNum];
            for (int i = 0; i < capacity; i++)
            {
                for (int j = 0; j < sector; j++)
                {
                    for (int k = 0; k < trainNum; i++)
                    {
                        capacityMatrix[i, j, k] = 1;
                    }
                }
            }


        }

        public void InitialProduct(double discountRatio, double theta, int OD_pair, double OD_CertainDemand, double OD_FullTicketPrice)
        {
            int temp = 0;
            for (int i = 0; i < OD_pair; i++)
            {

                OD_Dimension oD_Dimension = new OD_Dimension();
                oD_Dimension.OD_ID = temp;
                oD_Dimension.OD_FullTicketPrice = OD_FullTicketPrice;


            }
        }
    }
}
