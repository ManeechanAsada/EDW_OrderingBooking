using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Avantik.Web.Service.Helpers
{
    public static class Date
    {
        public static string GetDateString(this DateTime dt)
        {
            return string.Format("{0:yyyyMMdd}", dt);
        }

        public static string GetDateTimeString(this DateTime dt)
        {
            return string.Format("{0:yyyyMMdd HH:mm:ss}", dt);
        }

        public static string GetDateDash(this DateTime dt)
        {
            return string.Format("{0:yyyy-MM-dd}", dt);
        }

        public static long DateDiffMinute(DateTime firstDate, DateTime secondDate)
        {
            return DateDiff(firstDate, secondDate, "M");
        }
        public static long DateDiffDay(DateTime firstDate, DateTime secondDate)
        {
            return DateDiff(firstDate, secondDate, "D");
        }

        #region Helper
        private static long DateDiff(DateTime firstDate, DateTime secondDate, string type)
        {
            TimeSpan span = (secondDate - firstDate);

            if (type == "D")
            {
                return span.Days;
            }
            else if (type == "H")
            {
                return span.Hours;
            }
            else if (type == "M")
            {
                return span.Minutes;
            }
            else if (type == "S")
            {
                return span.Seconds;
            }
            else
            {
                return 0;
            }
        }
        #endregion

        #region String
        public static bool HasSpecialCharacters(string str)
        {
            string specialCharacters = @"%!@#$%^&*()?/>.<,:;'\|}]{[_~`+=-" + "\"";
            char[] specialCharactersArray = specialCharacters.ToCharArray();

            int index = str.IndexOfAny(specialCharactersArray);
            //index == -1 no special characters
            if (index == -1)
                return false;
            else
                return true;
        }
        #endregion
    }
}
