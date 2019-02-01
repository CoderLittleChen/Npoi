using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace _01在指定目录下创建文件夹
{
    class Program
    {
        static void Main(string[] args)
        {
            string appPath = Environment.CurrentDirectory;
            string appendPath = "ExportDetailsDataToExcel";
            string newPath = Path.Combine(appPath, appendPath);
            Directory.CreateDirectory(newPath);
            Console.WriteLine(newPath);
            Console.ReadKey();
        }
    }
}
