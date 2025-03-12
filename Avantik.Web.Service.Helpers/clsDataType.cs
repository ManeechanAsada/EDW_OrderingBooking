using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Avantik.Web.Service.Helpers
{
    public static class DataType
    {
        public static string GetStringValue(object Value)
        {
            string StrValue = "";

            if (Value == null)
            {
                StrValue = "";
            }
            else
            {
                if (Value == System.DBNull.Value)
                {
                    StrValue = "";
                }
                else
                {
                    StrValue = Value.ToString().Trim();
                }
            }

            return StrValue;
        }

        public static bool IsStringPopulated(string InputString)
        {
            if (InputString == null)
                InputString = "";

            return (InputString.Trim().Length > 0);
        }

        public static bool IsGuid(string Input)
        {
            bool Result = false;

            if (IsStringPopulated(Input))
            {
                try
                {
                    Guid TheGuid = new Guid(Input.Trim());
                    Result = true;
                }
                catch { }
            }
            return Result;
        }

        public static bool IsDate(string Input)
        {
            bool Result = false;

            if (IsStringPopulated(Input))
            {
                try
                {
                    Convert.ToDateTime(Input);
                    Result = true;
                }
                catch { }
            }
            return Result;
        }

        public static bool IsXml(string Input)
        {
            bool Result = false;

            if (IsStringPopulated(Input))
            {
                try
                {
                    System.Xml.XmlDocument XmlDoc = new System.Xml.XmlDocument();
                    XmlDoc.LoadXml(Input);
                    Result = true;
                }
                catch { }
            }
            return Result;
        }

        public static bool IsDecimal(string Input)
        {
            bool Result = false;

            if (IsStringPopulated(Input))
            {
                try
                {
                    decimal theDecimal = Convert.ToDecimal(Input);
                    Result = true;
                }
                catch { }
            }
            return Result;
        }

        public static bool IsDouble(string Input)
        {
            bool Result = false;

            if (IsStringPopulated(Input))
            {
                try
                {
                    double theNumeric = Convert.ToDouble(Input.Trim());
                    Result = true;
                }
                catch { }
            }
            return Result;
        }

        public static bool IsBool(string Input)
        {
            bool Result = false;

            if (IsStringPopulated(Input))
            {
                try
                {
                    bool Temp = Convert.ToBoolean(Input.Trim());
                    Result = true;
                }
                catch
                {

                }
            }

            return Result;
        }

        public static bool IsInt16(string Input)
        {
            bool Result = false;

            if (IsStringPopulated(Input))
            {
                try
                {
                    Int16 theInt = Convert.ToInt16(Input.Trim());
                    Result = true;
                }
                catch { }
            }

            return Result;
        }

        public static bool IsInt32(string Input)
        {
            bool Result = false;

            if (IsStringPopulated(Input))
            {
                try
                {
                    int theInt = Convert.ToInt32(Input.Trim());
                    Result = true;
                }
                catch { }
            }

            return Result;
        }

        public static string RemoveWhiteSpace(string TheString)
        {
            TheString = GetStringValue(TheString).Trim();
            while (TheString.Contains(" "))
            {
                TheString = TheString.Replace(" ", "");
            }

            return TheString;
        }
    }
}
