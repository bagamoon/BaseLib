using log4net;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;

namespace LibCommon.Util
{
    /// <summary>
    /// 使用smtp server, 透過config設定值或自帶值寄送Email
    /// Config Key表:
    /// UL_Environment     系統環境(必填)    P:正式  T:測試(會於mail標題前註[測試環境])
    /// UL_SmtpHost        smtp host(必填)
    /// UL_SmtpUserName    smtp 帳號(不需帳號密碼請填空白)
    /// UL_SmtpPwd         smtp 密碼(不需帳號密碼請填空白)
    /// UL_MailSender      mail 寄件者(必填)
    /// UL_MailReceivers   mail 收件者(必填)
    /// UL_MailSubject     mail 標題
    /// UL_MailEnable      是否啟用mail寄送(必填)  Y:啟用  N:不啟用
    /// </summary>
    public static class EmailUtilities
    {
        private static ILog logger = LogManager.GetLogger(typeof(EmailUtilities).Name);

        #region "Config Key"

        //系統環境
        private static readonly string environmentKey = "UL_Environment";

        //smtp host
        private static readonly string smtpHostKey = "UL_SmtpHost";

        //smtp帳號
        private static readonly string userNameKey = "UL_SmtpUserName";

        //smtp密碼
        private static readonly string passwordKey = "UL_SmtpPwd";

        //email寄件者
        private static readonly string senderKey = "UL_MailSender";

        //email收件者
        private static readonly string receiversKey = "UL_MailReceivers";

        //email標題
        private static readonly string subjectKey = "UL_MailSubject";

        //是否啟用mail寄送
        private static readonly string mailEnableKey = "UL_MailEnable";

        //MailLog資料庫連線字串
        private static readonly string mailLogDBKey = "UL_MailLogDB";

        //測試環境在標題前加註的字串
        private static readonly string subjectPrefix = "[測試環境]";

        //測試環境在內容前加註的字串
        private static readonly string contentAffix = "============ 此通知為測試環境 ============";
        #endregion

        /// <summary>
        /// 透過config設定值或自帶值寄送Email
        /// </summary>
        /// <param name="content">信件內容</param>
        public static void SendMail(string content)
        {
            string smtpHost = ConfigurationManager.AppSettings[smtpHostKey];
            string userName = ConfigurationManager.AppSettings[userNameKey];
            string password = ConfigurationManager.AppSettings[passwordKey];
            string sender = ConfigurationManager.AppSettings[senderKey];
            string receivers = ConfigurationManager.AppSettings[receiversKey];
            string subject = ConfigurationManager.AppSettings[subjectKey];

            SendMail(smtpHost, userName, password, sender, receivers, subject, content, string.Empty);
        }

        /// <summary>
        /// 透過config設定值或自帶值寄送Email
        /// </summary>
        /// <param name="content">信件內容</param>
        /// <param name="attachedFilePaths">附檔路徑, 多個以";"分隔, 不需附件可放null</param>
        public static void SendMail(string content, string attachedFilePaths)
        {
            string smtpHost = ConfigurationManager.AppSettings[smtpHostKey];
            string userName = ConfigurationManager.AppSettings[userNameKey];
            string password = ConfigurationManager.AppSettings[passwordKey];
            string sender = ConfigurationManager.AppSettings[senderKey];
            string receivers = ConfigurationManager.AppSettings[receiversKey];
            string subject = ConfigurationManager.AppSettings[subjectKey];

            SendMail(smtpHost, userName, password, sender, receivers, subject, content, attachedFilePaths);
        }

        /// <summary>
        /// 透過config設定值或自帶值寄送Email
        /// </summary>
        /// <param name="subject">信件標題</param>
        /// <param name="content">信件內容</param>
        /// <param name="attachedFilePaths">附檔路徑, 多個以";"分隔, 不需附件可放null</param>
        public static void SendMail(string subject, string content, string attachedFilePaths)
        {
            string smtpHost = ConfigurationManager.AppSettings[smtpHostKey];
            string userName = ConfigurationManager.AppSettings[userNameKey];
            string password = ConfigurationManager.AppSettings[passwordKey];
            string sender = ConfigurationManager.AppSettings[senderKey];
            string receivers = ConfigurationManager.AppSettings[receiversKey];

            SendMail(smtpHost, userName, password, sender, receivers, subject, content, attachedFilePaths);
        }

        /// <summary>
        /// 透過config設定值或自帶值寄送Email
        /// </summary>        
        /// <param name="receivers">收件者, 多個以";"分隔</param>
        /// <param name="subject">信件標題</param>
        /// <param name="content">信件內容</param>
        /// <param name="attachedFilePaths">附檔路徑, 多個以";"分隔, 不需附件可放null</param>
        public static void SendMail(string receivers, string subject, string content, string attachedFilePaths)
        {
            string smtpHost = ConfigurationManager.AppSettings[smtpHostKey];
            string userName = ConfigurationManager.AppSettings[userNameKey];
            string password = ConfigurationManager.AppSettings[passwordKey];
            string sender = ConfigurationManager.AppSettings[senderKey];

            SendMail(smtpHost, userName, password, sender, receivers, subject, content, attachedFilePaths);
        }

        /// <summary>
        /// 透過config設定值或自帶值寄送Email
        /// </summary>
        /// <param name="sender">寄件者</param>
        /// <param name="receivers">收件者, 多個以";"分隔</param>
        /// <param name="subject">信件標題</param>
        /// <param name="content">信件內容</param>
        /// <param name="attachedFilePaths">附檔路徑, 多個以";"分隔, 不需附件可放null</param>
        public static void SendMail(string sender, string receivers, string subject, string content, string attachedFilePaths)
        {
            string smtpHost = ConfigurationManager.AppSettings[smtpHostKey];
            string userName = ConfigurationManager.AppSettings[userNameKey];
            string password = ConfigurationManager.AppSettings[passwordKey];

            SendMail(smtpHost, userName, password, sender, receivers, subject, content, attachedFilePaths);
        }

        /// <summary>
        /// 透過config設定值或自帶值寄送Email
        /// </summary>
        /// <param name="smtpHost">smtp host</param>
        /// <param name="userName">smtp帳號</param>
        /// <param name="password">smtp密碼</param>
        /// <param name="sender">寄件者</param>
        /// <param name="receivers">收件者, 多個以";"分隔</param>
        /// <param name="subject">信件標題</param>
        /// <param name="content">信件內容</param>        
        public static void SendMail(string smtpHost, string userName, string password, string sender, string receivers, string subject, string content)
        {
            SendMail(smtpHost, userName, password, sender, receivers, subject, content, string.Empty);
        }

        /// <summary>
        /// 透過config設定值或自帶值寄送Email
        /// </summary>
        /// <param name="smtpHost">smtp host</param>
        /// <param name="userName">smtp帳號</param>
        /// <param name="password">smtp密碼</param>
        /// <param name="sender">寄件者</param>
        /// <param name="receivers">收件者, 多個以";"分隔</param>
        /// <param name="subject">信件標題</param>
        /// <param name="content">信件內容</param>        
        /// <param name="attachedFilePaths">附檔路徑, 多個以";"分隔, 不需附件可放null</param>
        public static void SendMail(string smtpHost, string userName, string password, string sender, string receivers, string subject, string content, string attachedFilePaths)
        {
            IList<string> receiverList = default(IList<string>);

            if (receivers != null)
            {
                receiverList = receivers.Split(new string[] { ";" }, StringSplitOptions.RemoveEmptyEntries);
            }
            else
            {
                receiverList = new List<string>();
            }

            IList<string> attachedFilePathList = default(IList<string>);

            if (attachedFilePaths != null)
            {
                attachedFilePathList = attachedFilePaths.Split(new string[] { ";" }, StringSplitOptions.RemoveEmptyEntries);
            }
            else
            {
                attachedFilePathList = new List<string>();
            }

            SendMail(smtpHost, userName, password, sender, receiverList, subject, content, attachedFilePathList);
        }

        /// <summary>
        /// 透過config設定值或自帶值寄送Email
        /// </summary>
        /// <param name="smtpHost">smtp host</param>
        /// <param name="userName">smtp帳號</param>
        /// <param name="password">smtp密碼</param>
        /// <param name="sender">寄件者</param>
        /// <param name="receiverList">收件者List</param>
        /// <param name="subject">信件標題</param>
        /// <param name="content">信件內容</param>   
        /// <param name="attachedFilePathList">附檔路徑List</param>
        public static void SendMail(string smtpHost, string userName, string password, string sender, IList<string> receiverList, string subject, string content, IList<string> attachedFilePathList)
        {
            string environment = ConfigurationManager.AppSettings[environmentKey];
            string mailEnable = ConfigurationManager.AppSettings[mailEnableKey];
            string mailLogDB = ConfigurationManager.AppSettings[mailLogDBKey];

            if (smtpHost == null)
                throw new Exception("SmtpHost未設定, 請檢查config或傳入參數");

            if (userName == null)
                throw new Exception("UserName未設定, 請檢查config或傳入參數");

            if (password == null)
                throw new Exception("Password未設定, 請檢查config或傳入參數");

            if (sender == null)
                throw new Exception("Sender未設定, 請檢查config或傳入參數");

            if (subject == null)
                throw new Exception("Subject未設定, 請檢查config或傳入參數");

            if (content == null)
                throw new Exception("Content未設定, 請檢查config或傳入參數");

            if (environment == null)
                throw new Exception("Environment未設定, 請檢查config或傳入參數");

            if (mailEnable == null)
                throw new Exception("MailEnable未設定, 請檢查config或傳入參數");


            //啟用時才寄送mail
            if (mailEnable.ToUpper() == "Y" && receiverList.Count > 0)
            {
                //smtp設定
                SmtpClient smtp = new SmtpClient(smtpHost);
                if (string.IsNullOrWhiteSpace(userName) == false)
                {
                    smtp.Credentials = new NetworkCredential(userName, password);
                }

                string errorMessage = "";

                try
                {
                    logger.DebugFormat("寄送郵件, 寄件者: {0}, 收件者: {1}, 標題: {2}, 內容: {3}, 附件: {4}",
                                        sender,
                                        string.Join(";", receiverList),
                                        subject,
                                        content,
                                        string.Join(";", attachedFilePathList));

                    using (MailMessage mail = new MailMessage())
                    {

                        //寄件者
                        mail.From = new MailAddress(sender);

                        //若為測試環境則在標題及內容標註測試環境
                        if (environment.ToUpper() == "T")
                        {
                            subject = subjectPrefix + subject;
                            content = string.Format("{0} <br /><br /> {1} <br /><br /> {0}", contentAffix, content);
                        }

                        //標題
                        mail.Subject = subject;

                        //內容
                        AlternateView alt = AlternateView.CreateAlternateViewFromString(content, null, "text/html");
                        mail.AlternateViews.Add(alt);

                        //收件者list
                        foreach (string receiver in receiverList)
                        {
                            if (string.IsNullOrWhiteSpace(receiver) == false)
                            {
                                mail.To.Add(receiver.Trim());
                            }
                        }

                        //附檔list            
                        foreach (string filePath in attachedFilePathList)
                        {
                            if (string.IsNullOrWhiteSpace(filePath) == false && File.Exists(filePath) == true)
                            {
                                mail.Attachments.Add(new Attachment(filePath));
                            }
                        }

                        smtp.Send(mail);
                    }

                    logger.Debug("寄送完成");
                }
                catch (Exception ex)
                {
                    logger.Fatal(ex.Message, ex);
                    errorMessage = ex.Message;
                }
            }
        }
    }
}
