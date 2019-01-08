using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace _2Substring
{
    class Program
    {
        static void Main(string[] args)
        {
            string name = "自定义@@个人考勤汇总表明细";
            int index = name.LastIndexOf("@");

            name = name.Substring(index + 1);


            Console.WriteLine(name);
            Console.ReadKey();


        }
    }
}
