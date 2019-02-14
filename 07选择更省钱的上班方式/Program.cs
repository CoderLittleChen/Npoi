using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace _07选择更省钱的上班方式
{
    class Program
    {
        static void Main(string[] args)
        {
            //现有三种上班方式   判断哪种更省钱  按照每月22个工作日来计算
            //地铁打折方式：当月累计消费满100，8折；满150  5折   超出400  不打折   单程初始值设为5元


            //1、上下班都做地铁
            //2、上班做地铁    下班做公交

            //第一种方式 
            double sumMoneyInFirstPlan = 0;
            int eightDiscountStartIndexInFirstPlan = 0;
            int fiveDiscountStartIndeInFirstPlan = 0;
            for (int i = 0; i < 44; i++)
            {
                if (sumMoneyInFirstPlan >= 100 && sumMoneyInFirstPlan <= 150)
                {
                    if (eightDiscountStartIndexInFirstPlan == 0)
                    {
                        eightDiscountStartIndexInFirstPlan = i + 1;
                    }
                    sumMoneyInFirstPlan += 5 * 0.8;
                }
                else if (sumMoneyInFirstPlan >= 150 && sumMoneyInFirstPlan <= 400)
                {
                    if (fiveDiscountStartIndeInFirstPlan == 0)
                    {
                        fiveDiscountStartIndeInFirstPlan = i + 1;
                    }
                    sumMoneyInFirstPlan += 5 * 0.5;
                }
                else
                {
                    sumMoneyInFirstPlan += 5;
                }
            }


            //第二种方式
            double sumSubwayMoneyInSecondPlan = 0;
            double sumBusMoneyInSecondPlan = 0;
            int eightDiscountStartIndexInSecondPlan = 0;
            int fiveDiscountStartIndexInSecondPlan = 0;
            for (int i = 0; i < 22; i++)
            {
                if (sumSubwayMoneyInSecondPlan >= 100 && sumSubwayMoneyInSecondPlan <= 150)
                {
                    if (eightDiscountStartIndexInSecondPlan == 0)
                    {
                        eightDiscountStartIndexInSecondPlan = i + 1;
                    }
                    sumSubwayMoneyInSecondPlan += 5 * 0.8;
                }
                else if (sumSubwayMoneyInSecondPlan >= 150 && sumSubwayMoneyInSecondPlan <= 400)
                {
                    if (fiveDiscountStartIndexInSecondPlan == 0)
                    {
                        fiveDiscountStartIndexInSecondPlan = i + 1;
                    }
                    sumSubwayMoneyInSecondPlan += 5 * 0.5;
                }
                else
                {
                    sumSubwayMoneyInSecondPlan += 5;
                }
                sumBusMoneyInSecondPlan += 1.5;
            }

            if (eightDiscountStartIndexInFirstPlan % 2 != 0)
            {
                eightDiscountStartIndexInFirstPlan++;
            }
            if (fiveDiscountStartIndeInFirstPlan % 2 != 0)
            {
                fiveDiscountStartIndeInFirstPlan++;
            }


            Console.WriteLine(string.Format("第一种方式,第{0}个工作日开始8折", eightDiscountStartIndexInFirstPlan / 2));
            Console.WriteLine(string.Format("第一种方式,第{0}个工作日开始5折", fiveDiscountStartIndeInFirstPlan / 2));
            Console.WriteLine(string.Format("第一种方式月消费：{0}元", sumMoneyInFirstPlan));
            Console.WriteLine("--------------------");

            Console.WriteLine(string.Format("第一种方式,第{0}个工作日开始8折", eightDiscountStartIndexInSecondPlan));
            Console.WriteLine(string.Format("第一种方式,第{0}个工作日开始5折", fiveDiscountStartIndexInSecondPlan));
            Console.WriteLine(string.Format("第二种方式月消费：{0}元", sumBusMoneyInSecondPlan + sumSubwayMoneyInSecondPlan));


            Console.ReadKey();

        }
    }
}
