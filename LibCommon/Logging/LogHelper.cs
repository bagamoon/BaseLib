using log4net;
using log4net.Appender;
using log4net.Repository.Hierarchy;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;

namespace LibCommon.Logging
{
    /// <summary>
    /// 使用Log4Net記錄Log, 並提供method執行時間監控API及訊息視窗API
    /// </summary>
    public static class LogHelper
    {
        private static readonly TimeSpan bound;

        private static ILog logger = LogManager.GetLogger(typeof(LogHelper).Name);

        /// <summary>
        /// 效能監測Logger, LoggerName為PerformanceMonitor
        /// </summary>
        public static ILog PerformanceLogger { get; private set; }

        /// <summary>
        /// 訊息視窗Logger, LoggerName為MessageBoxDisplayer
        /// </summary>
        public static ILog MsgBoxLogger { get; private set; }

        /// <summary>
        /// 建構子, 初始化PerformanceLogger及MsgBoxLogger
        /// 初始化執行效能容忍時間上限(預設三秒), 讀取config: PerformanceBound
        /// </summary>
        static LogHelper()
        {
            PerformanceLogger = LogManager.GetLogger("PerformanceMonitor");
            MsgBoxLogger = LogManager.GetLogger("MessageBoxDisplayer");

            string boundStr = ConfigurationManager.AppSettings["PerformanceBound"];
            int boundInt = 0;
            if (int.TryParse(boundStr, out boundInt) == false)
            {
                boundInt = 3000;
            }
            bound = new TimeSpan(0, 0, 0, 0, boundInt);
        }

        /// <summary>
        /// 嘗試執行有回傳值的method, 若執行失敗會記錄log
        /// </summary>
        /// <typeparam name="TResult">回傳值的型別</typeparam>
        /// <param name="fun">欲執行的委派</param>
        /// <returns>執行結果</returns>
        public static ExecutedResult<TResult> TryExcute<TResult>(Expression<Func<TResult>> fun)
        {
            ExecutedResult<TResult> result;

            try
            {
                result = Excute(fun);
            }
            catch (Exception ex)
            {
                logger.Fatal(ex.Message, ex);
                result = new ExecutedResult<TResult> { IsSuccess = false, Exception = ex };
            }

            return result;
        }

        /// <summary>
        /// 嘗試執行無回傳值的method, 若執行失敗會記錄log
        /// </summary>
        /// <param name="act">欲執行的委派</param>
        /// <returns>是否執行成功</returns>
        public static bool TryExcute(Expression<Action> act)
        {
            bool isSuccess = false;

            try
            {
                Excute(act);
                isSuccess = true;
            }
            catch (Exception ex)
            {
                logger.Fatal(ex.Message, ex);
                isSuccess = false;
            }

            return isSuccess;
        }

        /// <summary>
        /// 執行有回傳值的method
        /// </summary>
        /// <typeparam name="TResult">回傳值的型別</typeparam>
        /// <param name="fun">欲執行的委派</param>
        /// <returns>執行結果</returns>
        public static ExecutedResult<TResult> Excute<TResult>(Expression<Func<TResult>> fun)
        {
            MethodCallExpression method = fun.Body as MethodCallExpression;
            string log = GetExpressionLog(method);

            DateTime start = DateTime.Now;
            TResult result = fun.Compile()();
            LogPerformance(DateTime.Now - start, log);

            ExecutedResult<TResult> executedResult = new ExecutedResult<TResult> { IsSuccess = true, Result = result };
            return executedResult;
        }

        /// <summary>
        /// 執行無回傳值的method
        /// </summary>
        /// <param name="act">欲執行的委派</param>
        public static void Excute(Expression<Action> act)
        {
            MethodCallExpression method = act.Body as MethodCallExpression;
            string log = GetExpressionLog(method);

            DateTime start = DateTime.Now;
            act.Compile()();
            LogPerformance(DateTime.Now - start, log);
        }

        /// <summary>
        /// 依據執行時間是否超過容忍上限記錄log
        /// </summary>
        /// <param name="timeCost">執行時間</param>
        /// <param name="logInfo">log訊息</param>
        private static void LogPerformance(TimeSpan timeCost, string logInfo)
        {
            if (timeCost < bound)
            {
                PerformanceLogger.Info(string.Format("{0} [{1} ms]", logInfo, timeCost.TotalMilliseconds));
            }
            else
            {
                PerformanceLogger.Warn(string.Format("{0} [{1} ms]", logInfo, timeCost.TotalMilliseconds));
            }
        }

        /// <summary>
        /// 以非同步方式刪除舊Log, 讀取KeepLogFileDays取得保留天數, 預設30天
        /// </summary>
        public static void DumpOldLogFiles()
        {
            int days = 30;

            string daysString = ConfigurationManager.AppSettings["KeepLogFileDays"];
            if (daysString != null)
            {
                int.TryParse(daysString, out days);
            }

            DumpOldLogFiles(days * -1);
        }

        /// <summary>
        /// 以非同步方式刪除舊Log, 傳入保留天數
        /// </summary>
        /// <param name="days"></param>
        public static void DumpOldLogFiles(int days)
        {
            logger.InfoFormat("清除歷史Log檔案, 篩選條件為小於現在時間 {0} 天", days);

            Task.Factory.StartNew(() =>
            {
                var rootAppender = ((Hierarchy)LogManager.GetRepository())
                                                         .GetAppenders().OfType<RollingFileAppender>()
                                                         .FirstOrDefault();

                string filePath = rootAppender != null ? rootAppender.File : string.Empty;

                if (string.IsNullOrWhiteSpace(filePath) == false && File.Exists(filePath))
                {
                    string dirPath = Path.GetDirectoryName(filePath);
                    if (Directory.Exists(dirPath))
                    {
                        DirectoryInfo dir = new DirectoryInfo(dirPath);

                        //取得最後更新時間小於保留日期前的log並刪除
                        IEnumerable<FileInfo> fileList = dir.GetFiles("*.log", SearchOption.AllDirectories)
                                                            .Where(p => p.LastWriteTime < DateTime.Today.AddDays(days));

                        foreach (FileInfo file in fileList)
                        {
                            try
                            {
                                logger.Info(string.Format("刪除檔案: {0}", file.FullName));
                                file.Delete();
                            }
                            catch (Exception ex)
                            {
                                logger.Fatal(string.Format("無法刪除檔案: {0}, Message: {1}", file.FullName, ex.Message), ex);
                            }
                        }
                    }
                }
            });
        }

        /// <summary>
        /// 取得委派的執行method資訊, 回傳"method: class.Method()"格式的字串
        /// </summary>
        /// <param name="methodCallExpression">委派內容資訊</param>
        /// <returns>"method: class.Method()"格式的字串</returns>
        private static string GetExpressionLog(MethodCallExpression methodCallExpression)
        {
            string log = "";
            if (methodCallExpression != null)
            {
                log = string.Format("method: {0}.{1}()", methodCallExpression.Method.ReflectedType.FullName, methodCallExpression.Method.Name);
            }

            return log;
        }
    }

    /// <summary>
    /// 執行結果
    /// </summary>
    /// <typeparam name="TResult">委派回傳值的型別</typeparam>
    public class ExecutedResult<TResult>
    {
        /// <summary>
        /// 是否執行成功
        /// </summary>
        public bool IsSuccess { get; set; }

        /// <summary>
        /// 委派回傳值
        /// </summary>
        public TResult Result { get; set; }

        /// <summary>
        /// 執行exception
        /// </summary>
        public Exception Exception { get; set; }
    }
}
