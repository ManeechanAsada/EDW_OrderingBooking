using System;
using System.Collections.Generic;
using System.Text;
using System.Configuration;

namespace Avantik.Web.Service.Helpers
{
    public static class ConfigHelper
    {
        public static string ToString(this System.Collections.Specialized.NameValueCollection setting, string strName)
        {
            if (setting != null && string.IsNullOrEmpty(setting[strName]) == false)
            {
                return setting[strName];
            }
            else
            {
                return string.Empty;
            }
        }

        public static byte ToByte(this System.Collections.Specialized.NameValueCollection setting, string strName)
        {
            if (setting != null && string.IsNullOrEmpty(setting[strName]) == false)
            {
                return Convert.ToByte(setting[strName]);
            }
            else
            {
                return 0;
            }
        }

        public static bool ToBoolean(this System.Collections.Specialized.NameValueCollection setting, string strName)
        {
            if (setting != null && string.IsNullOrEmpty(setting[strName]) == false)
            {
                return Convert.ToBoolean(setting[strName]);
            }
            else
            {
                return false;
            }
        }

        #region App Setting
        public static string ToString(string strName)
        {
            string strConfigValue = ConfigurationManager.AppSettings[strName];
            if (string.IsNullOrEmpty(strConfigValue) == false)
            {
                return strConfigValue;
            }
            else
            {
                return string.Empty;
            }
        }

        public static byte ToByte(string strName)
        {
            string strConfigValue = ConfigurationManager.AppSettings[strName];
            if (string.IsNullOrEmpty(strConfigValue) == false)
            {
                return Convert.ToByte(strConfigValue);
            }
            else
            {
                return 0;
            }
        }

        public static bool ToBoolean(string strName)
        {
            string strConfigValue = ConfigurationManager.AppSettings[strName];
            if (string.IsNullOrEmpty(strConfigValue) == false)
            {
                return Convert.ToBoolean(strConfigValue);
            }
            else
            {
                return false;
            }
        }
        #endregion
    }
}
