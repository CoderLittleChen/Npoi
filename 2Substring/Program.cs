using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace _02Substring
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            string name = "自定义@@个人考勤汇总表明细";
            int index = name.LastIndexOf("@");
            name = name.Substring(index + 1);
            Console.WriteLine(name);

            DateTime finalDate = new DateTime(2019, 1, 1, 0, 0, 0);
            finalDate = finalDate.AddDays(1 - finalDate.Day).AddMonths(1).AddDays(-1).Date;
            Console.WriteLine(finalDate);

            bool b1 = true;
            bool b2 = false;
            if (b1 | b2)
            {
                Console.WriteLine("111");
            }
            else
            {
                Console.WriteLine("222");
            }
            Console.WriteLine(DateTime.Now.ToString("yyyy-MM-dd"));

            Console.ReadKey();

        }
    }
}
