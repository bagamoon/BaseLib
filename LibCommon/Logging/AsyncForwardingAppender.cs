using log4net.Appender;
using log4net.Core;
using log4net.Util;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace LibCommon.Logging
{
    public class AsyncForwardingAppender : ForwardingAppender
    {
        private const int MAX_QUEUE_COUNT = 4096;
        private static int _asyncAppenderCount;
        private readonly Thread _loggingThread;
        private readonly BlockingCollection<LoggingEvent> _logEvents = new BlockingCollection<LoggingEvent>(MAX_QUEUE_COUNT);
        public AsyncForwardingAppender()
        {
            this._loggingThread = new Thread(new ThreadStart(this.LogThreadMethod))
            {
                IsBackground = true,
                Name = "AsyncForwardingAppender-" + Interlocked.Increment(ref AsyncForwardingAppender._asyncAppenderCount)
            };
            this._loggingThread.Start();
        }

        protected override void Append(LoggingEvent[] loggingEvents)
        {
            foreach (LoggingEvent loggingEvent in loggingEvents)
            {
                Append(loggingEvent);
            }
        }

        protected override void Append(LoggingEvent loggingEvent)
        {
            if (this._logEvents.Count < MAX_QUEUE_COUNT)
            {
                loggingEvent.Fix = FixFlags.ThreadName;
                this._logEvents.Add(loggingEvent);
                return;
            }
            LogLog.Warn(this.GetType(), "Queue is full, drop the overflow log message!");
        }
        private void LogThreadMethod()
        {
            while (true)
            {
                string appenderName = "";
                try
                {
                    LoggingEvent loggingEvent = this._logEvents.Take();
                    if (loggingEvent != null)
                    {
                        foreach (IAppender appender in this.Appenders)
                        {
                            appenderName = appender.Name;
                            appender.DoAppend(loggingEvent);
                        }
                    }
                    continue;
                }
                catch (ThreadAbortException)
                {
                    LogLog.Warn(this.GetType(), "The inner consuming thread is asked to be stopped.");
                }
                catch (ThreadInterruptedException)
                {
                    LogLog.Warn(this.GetType(), "The inner consuming thread is asked to be stopped.");
                }
                catch (Exception ex)
                {
                    LogLog.Warn(this.GetType(), string.Format("Failure in the independent thread-Appending within {0} : {1}", appenderName, ex));
                    Thread.Sleep(1000);
                    continue;
                }
                break;
            }
        }
        protected override void OnClose()
        {
            this._loggingThread.Abort();
            this._logEvents.Dispose();
            base.OnClose();
        }
    }
}
