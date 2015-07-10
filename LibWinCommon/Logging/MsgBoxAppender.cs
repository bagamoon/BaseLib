using log4net.Appender;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace LibWinCommon.Logging
{
    public class MsgBoxAppender : AppenderSkeleton
    {
        protected override void Append(log4net.Core.LoggingEvent loggingEvent)
        {
            string msg = base.RenderLoggingEvent(loggingEvent);

            switch (loggingEvent.Level.Name)
            {
                case "FATAL":
                case "ERROR":
                    MessageBox.Show(msg, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    break;

                case "WARN":
                    MessageBox.Show(msg, "警示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    break;

                case "INFO":
                    MessageBox.Show(msg, "資訊", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    break;

                case "DEBUG":
                    //debug不應顯示訊息
                    break;

            }
        }
    }
}
