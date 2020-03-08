using Quartz;
using Quartz.Impl;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ExpandComponents
{
    /********************************************************************************

   ** 类名称： TriggerManager

   ** 描述：触发器

   ** 引用： Quartz.dll，Common.Logging.dll,Common.Logging.Core.dll

   ** 配置 

    Global.asax配置

    protected void Application_Start()
     {
            TriggerManager<TestJob>.AddJob("Test", "0 0 0 * * ? *");//每天00:00:00更新
     }
   
    Web.config配置，配置在 <configSections> </configSections>节点下

   <!---Quartz启动也需要log4net-->
    <section name="Quartz" type="System.Configuration.NameValueSectionHandler, System, Version=2.5.0.0,Culture=neutral, PublicKeyToken=f6b8c98a402cc8a4"/>
    <sectionGroup name="common">
      <section name="logging" type="Common.Logging.ConfigurationSectionHandler, Common.Logging"/>
    </sectionGroup>

   ** 作者： LW

   *********************************************************************************/

    /// <summary>
    /// 触发器
    /// </summary> 
    public class TriggerManager<T> where T : class, IJob
    {
        //调度程序的工厂的接口中 实例一个具体的调度方法
        private static ISchedulerFactory schedulerFactory = new StdSchedulerFactory();
        //Job群组名
        private static string JOB_GROUP_NAME = "JOBGROUP_NAME";
        //触发器群
        private static string TRIGGER_GROUP_NAME = "TRIGGERGROUP_NAME";
        /// <summary>
        /// 添加一个定时任务，使用默认的任务组名，触发器名，触发器表达式
        /// </summary>
        /// <param name="pStrJobName">任务名</param>
        /// <param name="pStrCronExpress">触发器表达式，请参考Cron 表达式，表达式在线生成网址：http://www.bejson.com/othertools/cron/ 。例：0 0 0 * * ?（每天00:00:00开始执行） </param>
        public static void AddJob(string pStrJobName, string pStrCronExpress)
        {
            try
            {
                //接口中获取调度工厂的  GetScheduler()  方法
                IScheduler sched = schedulerFactory.GetScheduler();
                //创建任务
                IJobDetail job = JobBuilder.Create<T>().WithIdentity(pStrJobName, JOB_GROUP_NAME).Build();
                //创建触发器
                ITrigger trigger = TriggerBuilder.Create()
                    .WithIdentity(pStrJobName, TRIGGER_GROUP_NAME)
                    .WithCronSchedule(pStrCronExpress)
                    .Build();


                sched.ScheduleJob(job, trigger);
                sched.Start();
            }
            catch (Exception e)
            {

                throw new Exception(e.Message);
            }
        }
        /// <summary>
        /// 添加一个定时任务，间隔时长单位（小时）
        /// </summary>
        /// <param name="pStrJobName">任务名</param>
        /// <param name="hours">间隔小时</param>
        public static void AddJobIntervalHours(string pStrJobName, int hours)
        {
            try
            {
                //接口中获取调度工厂的  GetScheduler()  方法
                IScheduler sched = schedulerFactory.GetScheduler();
                //创建任务
                IJobDetail job = JobBuilder.Create<T>().WithIdentity(pStrJobName, JOB_GROUP_NAME).Build();
                //创建触发器
                ITrigger trigger = TriggerBuilder.Create()//创建一个触发器
                    .WithIdentity(pStrJobName, TRIGGER_GROUP_NAME)
                    .StartNow()
                    .WithSimpleSchedule(b => b.WithIntervalInHours(hours)//每隔多少小时执行一次
                    .RepeatForever())//无限循环执行
                    .Build();

                sched.ScheduleJob(job, trigger);
                sched.Start();
            }
            catch (Exception e)
            {

                throw new Exception(e.Message);
            }
        }
        /// <summary>
        /// 添加一个定时任务，间隔时长单位（秒）
        /// </summary>
        /// <param name="pStrJobName"></param>
        /// <param name="seconds">间隔秒数</param>
        public static void AddJobIntervalSeconds(string pStrJobName, int seconds)
        {
            try
            {
                //接口中获取调度工厂的  GetScheduler()  方法
                IScheduler sched = schedulerFactory.GetScheduler();
                //创建任务
                IJobDetail job = JobBuilder.Create<T>().WithIdentity(pStrJobName, JOB_GROUP_NAME).Build();
                //创建触发器
                ITrigger trigger = TriggerBuilder.Create()//创建一个触发器
                    .WithIdentity(pStrJobName, TRIGGER_GROUP_NAME)
                    .StartNow()
                    .WithSimpleSchedule(b => b.WithIntervalInSeconds(seconds)//每隔多少秒执行一次
                    .RepeatForever())//无限循环执行
                    .Build();

                sched.ScheduleJob(job, trigger);
                sched.Start();
            }
            catch (Exception e)
            {

                throw new Exception(e.Message);
            }
        }

        /// <summary>
        /// 移除一个任务(使用默认的任务组名，触发器名，触发器组名)
        /// </summary>
        /// <param name="pStrJobName"></param>
        public static void RemoveJob(string pStrJobName)
        {
            try
            {
                IScheduler sched = schedulerFactory.GetScheduler();
                JobKey jobkey = new JobKey(pStrJobName);
                TriggerKey triggerKey = new TriggerKey(pStrJobName, TRIGGER_GROUP_NAME);

                //停止触发器
                sched.PauseTrigger(triggerKey);
                //移除触发器
                sched.UnscheduleJob(triggerKey);
                //删除任务
                sched.DeleteJob(jobkey);
            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }
        }


        /// <summary>
        /// 修改一个任务的触发时间(使用默认的任务组名，触发器名，触发器组名)
        /// </summary>
        /// <param name="pStrJobName">任务名</param>
        /// <param name="pStrCronExpress">触发器表达式</param>
        public static void ModifyJobTime(string pStrJobName, string pStrCronExpress, IDictionary<string, object> pDictionary)
        {
            try
            {
                IScheduler sched = schedulerFactory.GetScheduler();
                TriggerKey triggerKey = new TriggerKey(pStrJobName, TRIGGER_GROUP_NAME);
                ICronTrigger trigger = (ICronTrigger)sched.GetTrigger(triggerKey);
                if (trigger == null)
                {
                    return;
                }
                RemoveJob(pStrJobName);
                AddJob(pStrJobName, pStrCronExpress);
            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }
        }


        /// <summary>
        ///  开启所有定时任务
        /// </summary>
        public static void StartAllJobs()
        {
            try
            {
                IScheduler sched = schedulerFactory.GetScheduler();
                sched.Start();
            }
            catch (Exception e)
            {

                throw new Exception(e.Message);
            }
        }

        /// <summary>
        /// 关闭所有的定时任务
        /// </summary>
        public static void ShutdownAllJobs()
        {
            try
            {
                IScheduler sched = schedulerFactory.GetScheduler();
                if (!sched.IsShutdown)
                {
                    sched.Shutdown();
                }
            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }
        }

        /// <summary>
        /// 恢复所有的任务
        /// </summary>
        public static void ResumeAllJobs()
        {
            try
            {
                IScheduler sched = schedulerFactory.GetScheduler();
                if (!sched.IsShutdown)
                {
                    sched.ResumeAll();
                }
            }
            catch (Exception e)
            {

                throw new Exception(e.Message);
            }
        }

        /// <summary>
        /// 暂停所有的作业
        /// </summary>
        public static void PauseAllJobs()
        {
            try
            {
                IScheduler sched = schedulerFactory.GetScheduler();
                sched.PauseAll();
            }
            catch (Exception e)
            {

                throw new Exception(e.Message);
            }
        }
    }

    ///// <summary>
    ///// 创建一个执行任务
    ///// </summary>
    //public class TestJob : IJob
    //{
    //    public void Execute(IJobExecutionContext context)
    //    {
    //        try
    //        {

    //        }
    //        catch (Exception e)
    //        {
    //            throw new Exception(e.Message);
    //        }
    //    }
    //}
}