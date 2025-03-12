using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Data.SqlClient;

namespace Avantik.Web.Service.Helpers.Database
{
    public static class DataHelpers
    {
        #region DataRow Helper
        public static string DBToString(this DataRow dr, string strValue)
        {
            if (dr.Table.Columns.Contains(strValue) == true)
            {
                if (dr[strValue] == null || dr[strValue] is System.DBNull)
                {
                    return string.Empty;
                }
                else
                {
                    return dr[strValue].ToString();
                }
            }
            else
            {
                return string.Empty;
            }
        }
        public static byte DBToByte(this DataRow dr, string strValue)
        {
            if (dr.Table.Columns.Contains(strValue) == true)
            {
                if (dr[strValue] == null || dr[strValue] is System.DBNull)
                {
                    return 0;
                }
                else
                {
                    return Convert.ToByte(dr[strValue]);
                }
            }
            else
            {
                return 0;
            }
        }
        public static DateTime DBToDateTime(this DataRow dr, string strValue)
        {
            if (dr.Table.Columns.Contains(strValue) == true)
            {
                if (dr[strValue] == null || dr[strValue] is System.DBNull)
                {
                    return DateTime.MinValue;
                }
                else
                {
                    return Convert.ToDateTime(dr[strValue]);
                }
            }
            else
            {
                return DateTime.MinValue;
            }
        }
        public static Guid DBToGuid(this DataRow dr, string strValue)
        {
            if (dr.Table.Columns.Contains(strValue) == true)
            {
                if (dr[strValue] == null || dr[strValue] is System.DBNull)
                {
                    return Guid.Empty;
                }
                else
                {
                    return new Guid(dr[strValue].ToString());
                }
            }
            else
            {
                return Guid.Empty;
            }

        }
        public static bool DBToBoolean(this DataRow dr, string strValue)
        {
            if (dr.Table.Columns.Contains(strValue) == true)
            {
                if (dr[strValue] == null || dr[strValue] is System.DBNull)
                {
                    return false;
                }
                else
                {
                    if (Convert.ToByte(dr[strValue]) == 0)
                    {
                        return false;
                    }
                    else
                    {
                        return true;
                    }
                }
            }
            else
            {
                return false;
            }
        }
        public static Int16 DBToInt16(this DataRow dr, string strValue)
        {
            if (dr.Table.Columns.Contains(strValue) == true)
            {
                if (dr[strValue] == null || dr[strValue] is System.DBNull)
                {
                    return 0;
                }
                else
                {
                    return Convert.ToInt16(dr[strValue]);
                }
            }
            else
            {
                return 0;
            }
        }
        public static Int32 DBToInt32(this DataRow dr, string strValue)
        {
            if (dr.Table.Columns.Contains(strValue) == true)
            {
                if (dr[strValue] == null || dr[strValue] is System.DBNull)
                {
                    return 0;
                }
                else
                {
                    return Convert.ToInt32(dr[strValue]);
                }
            }
            else
            {
                return 0;
            }
        }
        public static Int64 DBToInt64(this DataRow dr, string strValue)
        {
            if (dr.Table.Columns.Contains(strValue) == true)
            {
                if (dr[strValue] == null || dr[strValue] is System.DBNull)
                {
                    return 0;
                }
                else
                {
                    return Convert.ToInt64(dr[strValue]);
                }
            }
            else
            {
                return 0;
            }
        }
        public static decimal DBToDecimal(this DataRow dr, string strValue)
        {
            if (dr.Table.Columns.Contains(strValue) == true)
            {
                if (dr[strValue] == null || dr[strValue] is System.DBNull)
                {
                    return 0;
                }
                else
                {
                    return Convert.ToDecimal(dr[strValue]);
                }
            }
            else
            {
                return 0;
            }
        }
        #endregion
        #region DataRow Helper
        public static string DBToString(this SqlDataReader dr, string strValue)
        {
            if (dr[strValue] == null || dr[strValue] is System.DBNull)
            {
                return string.Empty;
            }
            else
            {
                return dr[strValue].ToString();
            }
        }
        public static byte DBToByte(this SqlDataReader dr, string strValue)
        {
            if (dr[strValue] == null || dr[strValue] is System.DBNull)
            {
                return 0;
            }
            else
            {
                return Convert.ToByte(dr[strValue]);
            }
        }
        public static DateTime DBToDateTime(this SqlDataReader dr, string strValue)
        {
            if (dr[strValue] == null || dr[strValue] is System.DBNull)
            {
                return DateTime.MinValue;
            }
            else
            {
                return Convert.ToDateTime(dr[strValue]);
            }
        }
        public static Guid DBToGuid(this SqlDataReader dr, string strValue)
        {
            if (dr[strValue] == null || dr[strValue] is System.DBNull)
            {
                return Guid.Empty;
            }
            else
            {
                return new Guid(dr[strValue].ToString());
            }

        }
        public static bool DBToBoolean(this SqlDataReader dr, string strValue)
        {
            if (dr[strValue] == null || dr[strValue] is System.DBNull)
            {
                return false;
            }
            else
            {
                if (Convert.ToByte(dr[strValue]) == 0)
                {
                    return false;
                }
                else
                {
                    return true;
                }
            }
        }
        public static Int16 DBToInt16(this SqlDataReader dr, string strValue)
        {
            if (dr[strValue] == null || dr[strValue] is System.DBNull)
            {
                return 0;
            }
            else
            {
                return Convert.ToInt16(dr[strValue]);
            }
        }
        public static Int32 DBToInt32(this SqlDataReader dr, string strValue)
        {
            if (dr[strValue] == null || dr[strValue] is System.DBNull)
            {
                return 0;
            }
            else
            {
                return Convert.ToInt32(dr[strValue]);
            }
        }
        public static Int64 DBToInt64(this SqlDataReader dr, string strValue)
        {
            if (dr[strValue] == null || dr[strValue] is System.DBNull)
            {
                return 0;
            }
            else
            {
                return Convert.ToInt64(dr[strValue]);
            }
        }
        public static decimal DBToDecimal(this SqlDataReader dr, string strValue)
        {
            if (dr[strValue] == null || dr[strValue] is System.DBNull)
            {
                return 0;
            }
            else
            {
                return Convert.ToDecimal(dr[strValue]);
            }
        }
        #endregion
        #region Assign Field
        public static string SetGuid(this Guid value)
        {
            if (value == Guid.Empty)
            {
                return null;
            }
            else
            {
                return value.ToString();
            }
        }
        #endregion
    }
}
