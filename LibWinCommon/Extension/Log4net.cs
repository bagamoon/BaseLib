using LibCommon.Logging;
using log4net;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LibWinCommon.Extension
{
    public static partial class Extensions
    {
        /// <summary>
        /// 寫入Info Level的Log並顯示MsgBox
        /// </summary>
        /// <param name="logger">logger</param>
        /// <param name="msg">訊息</param>
        /// <param name="ex">exception</param>
        public static void InfoWithShowMsg(this ILog logger, object msg, Exception ex = null)
        {
            logger.Info(msg, ex);
            LogHelper.MsgBoxLogger.Info(msg);
        }

        /// <summary>
        /// 寫入Warn Level的Log並顯示MsgBox
        /// </summary>
        /// <param name="logger">logger</param>
        /// <param name="msg">訊息</param>
        /// <param name="ex">exception</param>
        public static void WarnWithShowMsg(this ILog logger, object msg, Exception ex = null)
        {
            logger.Warn(msg, ex);
            LogHelper.MsgBoxLogger.Warn(msg);
        }

        /// <summary>
        /// 寫入Error Level的Log並顯示MsgBox
        /// </summary>
        /// <param name="logger">logger</param>
        /// <param name="msg">訊息</param>
        /// <param name="ex">exception</param>
        public static void ErrorWithShowMsg(this ILog logger, object msg, Exception ex = null)
        {
            logger.Error(msg, ex);
            LogHelper.MsgBoxLogger.Error(msg);
        }

        /// <summary>
        /// 寫入Fatal Level的Log並顯示MsgBox
        /// </summary>
        /// <param name="logger">logger</param>
        /// <param name="msg">訊息</param>
        /// <param name="ex">exception</param>
        public static void FatalWithShowMsg(this ILog logger, object msg, Exception ex = null)
        {
            logger.Fatal(msg, ex);
            LogHelper.MsgBoxLogger.Fatal(msg);
        }
    }
}
