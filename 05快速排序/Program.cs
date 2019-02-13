using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace _05快速排序
{
    class Program
    {
        static void Main(string[] args)
        {
            int[] nums = { 11, 2, 4, 55, 1, 9, 8, 6 };
            QuickSortHelper quickSortHelper = new QuickSortHelper();
            nums= quickSortHelper.QuickSort(nums);
            for (int i = 0; i < nums.Length; i++)
            {
                Console.Write(nums[i]+"  ");
            }
            Console.ReadKey();

        }
    }
}
