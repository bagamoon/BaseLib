using log4net.Appender;
using log4net.Core;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace LibWinCommon.Logging
{
    public class TextBoxAppender : AppenderSkeleton
    {
        private RichTextBox _textBox;
        public string FormName { get; set; }
        public string TextBoxName { get; set; }

        protected override void Append(LoggingEvent loggingEvent)
        {
            if (_textBox == null)
            {
                if (String.IsNullOrEmpty(FormName) ||
                    String.IsNullOrEmpty(TextBoxName))
                    return;

                Form form = Application.OpenForms[FormName];
                if (form == null)
                    return;

                _textBox = FindControlRecursive(form, TextBoxName) as RichTextBox;
                if (_textBox == null)
                    return;

                form.FormClosing += (s, e) => _textBox = null;
            }

            if (_textBox.IsDisposed == false)
            {
                try
                {
                    if (_textBox.InvokeRequired)
                    {
                        _textBox.Invoke(
                            (MethodInvoker)(
                            () =>
                            {
                                AppendText(loggingEvent);
                            }));
                    }
                    else
                    {
                        AppendText(loggingEvent);
                    }
                }
                catch (ThreadInterruptedException)
                {
                    //因主程式視窗關閉造成的錯誤, 不處理
                }
            }
        }

        private void AppendText(LoggingEvent loggingEvent)
        {
            switch (loggingEvent.Level.Name)
            {
                case "FATAL":
                case "ERROR":
                    _textBox.SelectionColor = Color.Red;
                    break;

                case "WARN":
                    _textBox.SelectionColor = Color.Orange;
                    break;

                case "INFO":
                case "DEBUG":
                    _textBox.SelectionColor = Color.Black;
                    break;

            }

            _textBox.AppendText(base.RenderLoggingEvent(loggingEvent));
            _textBox.ScrollToCaret();
        }

        private Control FindControlRecursive(Control root, string name)
        {
            if (root.Name == name) return root;
            foreach (Control c in root.Controls)
            {
                Control t = FindControlRecursive(c, name);
                if (t != null) return t;
            }
            return null;
        }
    }
}
