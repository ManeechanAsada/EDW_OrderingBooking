using System;
using System.Collections.Generic;
using System.Text;
using System.Net.Mail;
using System.Configuration;
using System.Collections.Specialized;
using System.IO;
using System.Web;
using System.Reflection;
using System.Diagnostics;
namespace Avantik.Web.Service.Helpers
{
    public class Logger
    {
        public enum LogType
        {
            File,
            Mail
        };
     
        //private static declaration.
        static LogType _logType;
        static Logger _instance;
        static string _LogAddress;
        static string _smtpServer;
        static string _application;
        static string _airline;
        static bool _ExistingInstance;
        #region Property
        public static Logger Instance(LogType logType)
        {
            if (_instance == null)
            {
                _instance = new Logger();
                _ExistingInstance = false;

                //initial configuration.
                NameValueCollection setting= (NameValueCollection)ConfigurationManager.GetSection("ErrorLog");
                if (setting.ToString("logtype") == "MAIL")
                {
                    //Set log Type.
                    _logType = LogType.Mail;

                    //Set mail error address
                    _smtpServer = setting.ToString("SmtpServer");
                    _LogAddress = setting.ToString("ErrorTo");
                    _application = setting.ToString("application");
                    _airline = setting.ToString("airline");
                }
            }
            else
            {
                _ExistingInstance = true;
            }

            return _instance;
        }
        #endregion

        #region Method
        public void WriteLog(Exception ex, string inputParameter)
        {
            CreateLog(_application,
                    _airline,
                    inputParameter,
                    ex.Message,
                    ex.StackTrace,
                    ex.TargetSite.DeclaringType.FullName,
                    ex.TargetSite.Name,
                    string.Empty);
        }
        public void WriteLog(string inputParameter, string logMessage, string errorLocation, string functionName, string applicationName, string traceMessage)
        {
            CreateLog(_application,
                      _airline,
                      inputParameter,
                      logMessage,
                      traceMessage,
                      errorLocation,
                      functionName,
                      applicationName);
        }
        #endregion

        #region Helper
        private void CreateLog(string application,
                               string airline,
                               string strInput, 
                               string strMessage, 
                               string strTrace, 
                               string strLocation, 
                               string strFunctionName,
                               string ApplicationType)
        {
            try
            {
                if (_logType == LogType.Mail)
                {
                    WriteLogMail(application,
                                 airline,
                                 strInput,
                                 strMessage,
                                 strTrace,
                                 strLocation,
                                 strFunctionName,
                                 ApplicationType);
                }
                else
                {
                    WriteLogFile(strInput,
                                 strMessage,
                                 strTrace,
                                 strLocation,
                                 strFunctionName,
                                 ApplicationType);
                }
            }
            catch
            { }
        }

        private void WriteLogMail(string application,
                                  string airline,
                                  string strInput,
                                  string strMessage,
                                  string strTrace,
                                  string strLocation,
                                  string strFunctionName,
                                  string ApplicationType)
        {
            try
            {
                if (string.IsNullOrEmpty(_smtpServer) == false)
                {
                    //Implement Writelog File.
                    StringBuilder stbHtml = new StringBuilder();

                    stbHtml.Append("<html xmlns='http://www.w3.org/1999/xhtml'>");
                    stbHtml.Append("<body>");
                    stbHtml.Append("<table width='400' border='0' cellpadding='4' cellspacing='2' bgcolor='#B4CBD6' style='margin:0px auto; font-family:Arial, Helvetica, sans-serif; font-size:12px; color:#304974;'>");
                    stbHtml.Append("<tr>");
                    stbHtml.Append("<td width='20%' valign='top' bgcolor='#FFFFFF' style='font-weight:bold;'>Date</td>");
                    stbHtml.Append("<td valign='top' bgcolor='#FFFFFF'>" + string.Format("{0:yyyy-MM-dd HH:mm:ss}", DateTime.Now) + " [Existing Instance = " + _ExistingInstance.ToString() + "]</td>");
                    stbHtml.Append("</tr>");
                    stbHtml.Append("<tr>");
                    stbHtml.Append("<td width='20%' valign='top' bgcolor='#FFFFFF' style='font-weight:bold;'>App Type</td>");
                    stbHtml.Append("<td valign='top' bgcolor='#FFFFFF'>" + application + "</td>");
                    stbHtml.Append("</tr>");
                    stbHtml.Append("<tr>");
                    stbHtml.Append("<td width='20%' valign='top' bgcolor='#FFFFFF' style='font-weight:bold;'>Airline</td>");
                    stbHtml.Append("<td valign='top' bgcolor='#FFFFFF'>" + airline + "</td>");
                    stbHtml.Append("</tr>");
                    stbHtml.Append("<tr>");
                    stbHtml.Append("<td width='20%' valign='top' bgcolor='#FFFFFF' style='font-weight:bold;'>Location</td>");
                    stbHtml.Append("<td valign='top' bgcolor='#FFFFFF'>" + strLocation + "</td>");
                    stbHtml.Append("</tr>");
                    stbHtml.Append("<tr>");
                    stbHtml.Append("<td width='20%' valign='top' bgcolor='#FFFFFF' style='font-weight:bold;'>Function</td>");
                    stbHtml.Append("<td valign='top' bgcolor='#FFFFFF'>" + strFunctionName + "</td>");
                    stbHtml.Append("</tr>");
                    stbHtml.Append("<tr>");
                    stbHtml.Append("<td width='20%' valign='top' bgcolor='#FFFFFF' style='font-weight:bold;'>Input Parameter</td>");
                    stbHtml.Append("<td valign='top' bgcolor='#FFFFFF'>");
                    stbHtml.Append(strInput.Replace("<", "&lt;").Replace(">", "&gt;"));
                    stbHtml.Append("</td>");
                    stbHtml.Append("</tr>");
                    stbHtml.Append("<tr>");
                    stbHtml.Append("<td width='20%' valign='top' bgcolor='#FFFFFF' style='font-weight:bold;'>Massage</td>");
                    stbHtml.Append("<td valign='top' bgcolor='#FFFFFF'>");
                    stbHtml.Append(strMessage);
                    stbHtml.Append("</td>");
                    stbHtml.Append("</tr>");
                    stbHtml.Append("<tr>");
                    stbHtml.Append("<td width='20%' valign='top' bgcolor='#FFFFFF' style='font-weight:bold;'>Trace</td>");
                    stbHtml.Append("<td valign='top' bgcolor='#FFFFFF'>");
                    stbHtml.Append(strTrace);
                    stbHtml.Append("</td>");
                    stbHtml.Append("</tr>");
                    stbHtml.Append("</table>");
                    stbHtml.Append("</body>");
                    stbHtml.Append("</html>");

                    //Call Mail function
                    //Send Error Email.
                    SendMailSmtp(_smtpServer,
                                 System.Environment.MachineName + "@mercator.asia",
                                 _LogAddress,
                                 string.Empty,
                                 "[" + strFunctionName + "] - " + strLocation,
                                 stbHtml.ToString(),
                                 true);
                }
            }
            catch
            { }
        }

        private void WriteLogFile(string strInput,
                                  string strMessage,
                                  string strTrace,
                                  string strLocation,
                                  string strFunctionName,
                                  string ApplicationType)
        {

            //Implement Writelog File.
            StringBuilder stbHtml = new StringBuilder();

            stbHtml.Append("***********************Date****************************");
            stbHtml.Append(string.Format("{0:yyyy-MM-dd HH:mm:ss}", DateTime.Now) + Environment.NewLine);
            stbHtml.Append("Location :  " + strLocation + Environment.NewLine);
            stbHtml.Append("Function :  " + strFunctionName + Environment.NewLine);
            stbHtml.Append("Input Parameter----------------------------------------" + Environment.NewLine);
            stbHtml.Append(strInput + Environment.NewLine);
            stbHtml.Append("Massage------------------------------------------------" + Environment.NewLine);
            stbHtml.Append(strMessage + Environment.NewLine);
            stbHtml.Append("Trace--------------------------------------------------"  + Environment.NewLine);
            stbHtml.Append(strTrace +  Environment.NewLine);

            //Save log file.

        }

        public static void SaveLog(string strFunctionName, DateTime dtStart, DateTime dtEnd, string strErrorMessage, string strInput)
        {
            if (strErrorMessage.Length > 0)
            {
                NameValueCollection setting = (NameValueCollection)ConfigurationManager.GetSection("ErrorLog");

                string strPath = setting.ToString("LogPath") + @"\" + String.Format("{0:yyyyMMdd}", DateTime.Now) + ".log";
                StringBuilder stb = new StringBuilder();
                StreamWriter stw = null;
                try
                {
                    using (stw = new StreamWriter(strPath, true))
                    {
                        stb.Append("------------------------------(" + strFunctionName + ")" + Environment.NewLine);
                        stb.Append("******Start " + String.Format("{0:dd/MM/yyyy hh:mm:ss}", dtStart) + Environment.NewLine);
                        stb.Append("******End " + String.Format("{0:dd/MM/yyyy hh:mm:ss}", dtEnd) + Environment.NewLine);
                        stb.Append("******Message" + Environment.NewLine);
                        stb.Append(strErrorMessage + Environment.NewLine);
                        stb.Append("Input Data" + Environment.NewLine);
                        stb.Append(strInput + Environment.NewLine);

                        stw.WriteLine(stb.ToString());
                        stw.Flush();
                    }
                }
                catch
                {
                    if (stw != null)
                    {
                        stw.Close();
                    }
                }
            }
        }
        public static string GetDllInfo()
        {
            string result = string.Empty;
            NameValueCollection setting = (NameValueCollection)ConfigurationManager.GetSection("ErrorLog");
            string path = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase);
            path = path.Replace(@"file:\", "");
            string dllPath1 = setting.ToString("dllPath");
            string dllPath2 = path;

            string dll1 = dllPath1 + @"\tikAeroProcess.dll";
            string dll2 = dllPath2 + @"\Avantik.Web.Service.dll";

            FileInfo oFileInfo = new FileInfo(dll1);
            DateTime dt = oFileInfo.CreationTime;
            FileVersionInfo myFileVersionInfo = FileVersionInfo.GetVersionInfo(dll1);
            string FileVersion = myFileVersionInfo.FileVersion;
            string ProductVersion = myFileVersionInfo.ProductVersion;

            FileInfo oFileInfo2 = new FileInfo(dll2);
            DateTime dt2 = oFileInfo2.CreationTime;
            FileVersionInfo myFileVersionInfo2 = FileVersionInfo.GetVersionInfo(dll2);
            string FileVersion2 = myFileVersionInfo2.FileVersion;
            string ProductVersion2 = myFileVersionInfo2.ProductVersion;

            result = @"tikAeroProcess.dll "
                + Environment.NewLine +
                "FileVersion: " + FileVersion
                + Environment.NewLine +
                "ProductVersion: " + ProductVersion
                + Environment.NewLine +
                "CreationTime: " + dt
                + Environment.NewLine
                 + Environment.NewLine
                + "Avantik.Web.Service.dll"
                + Environment.NewLine +
                "FileVersion: " + FileVersion2
                + Environment.NewLine +
                "ProductVersion: " + ProductVersion2
                 + Environment.NewLine +
                "CreationTime: " + dt2;
            return result;
        }

        public static string GetLogModify(string path)
        {
            string result = string.Empty;
            NameValueCollection setting = (NameValueCollection)ConfigurationManager.GetSection("ErrorLog");
            string logPath = setting.ToString("LogPath");
            string filedPath = logPath + @"\" + path + ".log";
            string[] filePaths = Directory.GetFiles(logPath + @"\", "*.log");

            // if not path get latest log
            if (string.IsNullOrEmpty(path))
            {
                filedPath = filePaths[filePaths.Length - 1];
            }

            if (File.Exists(filedPath))
            {
                result = System.IO.File.ReadAllText(filedPath);
            }
            else
            {
                result = "File not found.";
            }

            return result;
        }


        public static string DeleteLog(string path)
        {
            string result = string.Empty;
            NameValueCollection setting = (NameValueCollection)ConfigurationManager.GetSection("ErrorLog");
            string logPath = setting.ToString("LogPath");
            string filedPath = logPath + @"\" + path + ".log";
            string[] filePaths = Directory.GetFiles(logPath + @"\", "*.log");

            try
            {
                if (!string.IsNullOrEmpty(path))
                {
                    for (int i = 0; i < filePaths.Length; i++)
                    {
                        if (filePaths[i].Contains(path))
                        {
                            if (File.Exists(filePaths[i]))
                            {
                                System.IO.File.Delete(filePaths[i]);
                            }

                        }
                    }
                }

                StringBuilder sb = new StringBuilder();
                string[] updatedFilePaths = Directory.GetFiles(logPath + @"\", "*.log");

                for (int i = 0; i < updatedFilePaths.Length; i++)
                {
                    sb.Append(updatedFilePaths[i]);
                    sb.Append(Environment.NewLine);
                }

                result = sb.ToString();
            }
            catch
            {

            }

            return result;
        }

        public static void SaveLogModify(DateTime dtStart, DateTime dtEnd, string strInput, string token)
        {
                NameValueCollection setting = (NameValueCollection)ConfigurationManager.GetSection("ErrorLog");

                string strPath = setting.ToString("LogPath") + @"\cob" +String.Format("{0:yyyyMMddtt}", DateTime.Now) + ".log";
                StringBuilder stb = new StringBuilder();
                StreamWriter stw = null;
                try
                {
                    using (stw = new StreamWriter(strPath, true))
                    {
                        stb.Append("------------------------------" + Environment.NewLine);
                        stb.Append("******Start " + String.Format("{0:dd/MM/yyyy hh:mm:ss}", dtStart) + Environment.NewLine);
                        stb.Append("******End " + String.Format("{0:dd/MM/yyyy hh:mm:ss}", dtEnd) + Environment.NewLine);
                        stb.Append("Input" + Environment.NewLine);
                        stb.Append(token + Environment.NewLine);
                        stb.Append(strInput + Environment.NewLine);

                        stw.WriteLine(stb.ToString());
                        stw.Flush();
                    }
                }
                catch
                {
                    if (stw != null)
                    {
                        stw.Close();
                    }
                }
        }

        private bool SendMailSmtp(string smtpServer, string mailForm, string mailTo, string mailBcc, string mailSubject, string mailBody, bool isHtml)
        {
            bool bResult = false;
            if (smtpServer.Length > 0)
            {
                MailMessage message = null;
                try
                {
                    // Command line argument must the the SMTP host.
                    SmtpClient client = new SmtpClient(smtpServer);
                    MailAddress from = new MailAddress(mailForm,
                                                       mailForm,
                                                       System.Text.Encoding.UTF8);

                    MailAddress to = new MailAddress(mailTo);
                    message = new MailMessage(from, to);

                    if (mailBcc.Length > 0)
                    { message.Bcc.Add(mailBcc); }

                    message.IsBodyHtml = isHtml;
                    message.Body = mailBody;

                    message.BodyEncoding = System.Text.Encoding.UTF8;
                    message.Subject = mailSubject;
                    message.SubjectEncoding = System.Text.Encoding.UTF8;

                    client.Send(message);
                    bResult = true;
                }
                catch
                {
                    bResult = false;
                }
                finally
                {
                    if (message != null)
                    {
                        // Clean up.
                        message.Dispose();
                    }
                }
            }
            else
            { bResult = false; }

            return bResult;
        }
        #endregion
    }
}
