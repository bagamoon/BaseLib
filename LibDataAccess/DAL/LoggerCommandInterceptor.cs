using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.Entity.Infrastructure.Interception;
using log4net;
using System.Data.Common;

namespace LibDataAccess.DAL
{
    /// <summary>
    /// 攔截Entity Framework的指令執行並記錄log
    /// 需先執行 DbInterception.Add(new LoggerCommandInterceptor()); 註冊此類別
    /// 整個應用程式只需註冊一次
    /// </summary>
    public class LoggerCommandInterceptor : IDbCommandInterceptor
    {
        private ILog logger = LogManager.GetLogger("CommandInterceptor");

        /// <summary>
        /// ExecuteNonQuery執行中
        /// </summary>
        /// <param name="command">資料庫指令資訊</param>
        /// <param name="interceptionContext">執行資訊</param>
        public void NonQueryExecuting(DbCommand command, DbCommandInterceptionContext<int> interceptionContext)
        {
            LogCommand(command, interceptionContext);
        }

        /// <summary>
        /// ExecuteNonQuery執行後
        /// </summary>
        /// <param name="command">資料庫指令資訊</param>
        /// <param name="interceptionContext">執行資訊</param>
        public void NonQueryExecuted(DbCommand command, DbCommandInterceptionContext<int> interceptionContext)
        {
            LogIfError(command, interceptionContext); 
        }

        /// <summary>
        /// ExecuteReader執行中
        /// </summary>
        /// <param name="command">資料庫指令資訊</param>
        /// <param name="interceptionContext">執行資訊</param>
        public void ReaderExecuting(DbCommand command, DbCommandInterceptionContext<DbDataReader> interceptionContext)
        {
            LogCommand(command, interceptionContext);
        }

        /// <summary>
        /// ExecuteReader執行後
        /// </summary>
        /// <param name="command">資料庫指令資訊</param>
        /// <param name="interceptionContext">執行資訊</param>
        public void ReaderExecuted(DbCommand command, DbCommandInterceptionContext<DbDataReader> interceptionContext)
        {
            LogIfError(command, interceptionContext); 
        }

        /// <summary>
        /// ExecuteScalar執行中
        /// </summary>
        /// <param name="command">資料庫指令資訊</param>
        /// <param name="interceptionContext">執行資訊</param>
        public void ScalarExecuting(DbCommand command, DbCommandInterceptionContext<object> interceptionContext)
        {
            LogCommand(command, interceptionContext);
        }

        /// <summary>
        /// ExecuteScalar執行後
        /// </summary>
        /// <param name="command">資料庫指令資訊</param>
        /// <param name="interceptionContext">執行資訊</param>
        public void ScalarExecuted(DbCommand command, DbCommandInterceptionContext<object> interceptionContext)
        {
            LogIfError(command, interceptionContext); 
        }

        /// <summary>
        /// 記錄執行scripts及params
        /// </summary>
        /// <typeparam name="TResult">回傳結果類型</typeparam>
        /// <param name="command">資料庫指令資訊</param>
        /// <param name="interceptionContext">執行資訊</param>
        private void LogCommand<TResult>(DbCommand command, DbCommandInterceptionContext<TResult> interceptionContext)
        {
            if (command.Transaction != null)
            {
                logger.DebugFormat("Execute command: \r\n{0}, Parameters: \r\n{1} \r\nTransactionId: {2}, level: {3}",
                                  command.CommandText,
                                  TranslateParamsToString(command.Parameters),
                                  command.Transaction.GetHashCode(),
                                  command.Transaction.IsolationLevel);
            }
            else
            {
                logger.DebugFormat("Execute command: \r\n{0}, Parameters: \r\n{1}",
                                      command.CommandText,
                                      TranslateParamsToString(command.Parameters));
            }
        }

        /// <summary>
        /// 記錄發生錯誤的scripts及params
        /// </summary>
        /// <typeparam name="TResult">回傳結果類型</typeparam>
        /// <param name="command">資料庫指令資訊</param>
        /// <param name="interceptionContext">執行資訊</param>
        private void LogIfError<TResult>(DbCommand command, DbCommandInterceptionContext<TResult> interceptionContext)
        {
            if (interceptionContext.Exception != null)
            {
                if (command.Transaction != null)
                {
                    logger.ErrorFormat("Command: \r\n{0}, Parameters: \r\n{1} \r\ngot exception: \r\n{2} \r\nTransactionId: {3}",
                                       command.CommandText,
                                       TranslateParamsToString(command.Parameters),
                                       interceptionContext.Exception,
                                       command.Transaction.GetHashCode());
                }
                else
                {
                    logger.ErrorFormat("Command: \r\n{0}, Parameters: \r\n{1} \r\ngot exception: \r\n{2}",
                                           command.CommandText,
                                           TranslateParamsToString(command.Parameters),
                                           interceptionContext.Exception);
                }
            }
        }

        /// <summary>
        /// 將DbParameterCollection的每個參數名稱及值組為字串並以逗號區隔
        /// </summary>
        /// <param name="collection">DbParamter的集合</param>
        /// <returns></returns>
        private string TranslateParamsToString(DbParameterCollection collection)
        {
            string[] paramsArray = new string[collection.Count];

            for (int i = 0; i <= collection.Count - 1; i++)
            {
                paramsArray[i] = string.Format("{0} = {1}", collection[i].ParameterName, collection[i].Value);
            }

            return string.Join(", ", paramsArray);
        }
    }
}
