using Quartz;
using Quartz.Impl;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace _04Quarz.Net入门
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine(DateTime.Now.ToString("r"));
            //首先创建一个作业调度池
            ISchedulerFactory schedulerFactory = new StdSchedulerFactory();
            IScheduler scheduler = (IScheduler)schedulerFactory.GetScheduler();

            //创建出来一个具体的作业
            //IJobDetail jobDetail = JobBuilder.Create<JobDemo>().Build();



            Console.ReadKey();  
        }
    }

    public class JobDemo
    {
        //Task IJob.Execute(IJobExecutionContext context)
        //{
        //    //这里是作业调度每次定时执行方法    
        //    Console.WriteLine(DateTime.Now.ToString("r"));
        //}
    }

}
