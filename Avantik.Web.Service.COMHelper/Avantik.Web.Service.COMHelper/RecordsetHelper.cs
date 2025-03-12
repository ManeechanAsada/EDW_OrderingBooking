using System;
using System.Collections.Generic;
using System.Text;
using System.Reflection;
using System.Data;
using ADODB;

namespace Avantik.Web.Service.COMHelper
{
    public static class RecordsetHelper
    {
        public static string ToRsString(this Guid gValue)
        {
            if (gValue.Equals(Guid.Empty))
            {
                return null;
            }
            else
            {
                return "{" + gValue.ToString().ToUpper() + "}";
            }

        }
        public static string ToString(ADODB.Recordset rs, string strName)
        {

            if (ValidRsValue(rs, strName) == false)
            {
                return string.Empty;
            }
            else
            {
                return rs.Fields[strName].Value.ToString();
            }
        }
        public static byte ToByte(ADODB.Recordset rs, string strName)
        {
            if (ValidRsValue(rs, strName) == false)
            {
                return 0;
            }
            else
            {
                return Convert.ToByte(rs.Fields[strName].Value);
            }
        }
        public static bool ToBoolean(ADODB.Recordset rs, string strName)
        {
            if (ValidRsValue(rs, strName) == false)
            {
                return false;
            }
            else
            {
                return Convert.ToBoolean(rs.Fields[strName].Value);
            }
        }
        private static Char ToChar(ADODB.Recordset rs, string strName)
        {
            if (ValidRsValue(rs, strName) == false)
            {
                return Char.MinValue;
            }
            else
            {
                return Convert.ToChar(rs.Fields[strName].Value);
            }
        }
        public static Int16 ToInt16(ADODB.Recordset rs, string strName)
        {
            if (ValidRsValue(rs, strName) == false)
            {
                return 0;
            }
            else
            {
                return Convert.ToInt16(rs.Fields[strName].Value);
            }
        }
        public static Int32 ToInt32(ADODB.Recordset rs, string strName)
        {
            if (ValidRsValue(rs, strName) == false)
            {
                return 0;
            }
            else
            {
                return Convert.ToInt32(rs.Fields[strName].Value);
            }
        }
        public static Int64 ToInt64(ADODB.Recordset rs, string strName)
        {
            if (ValidRsValue(rs, strName) == false)
            {
                return 0;
            }
            else
            {
                return Convert.ToInt64(rs.Fields[strName].Value);
            }
        }
        private static Single ToSingle(ADODB.Recordset rs, string strName)
        {
            if (ValidRsValue(rs, strName) == false)
            {
                return 0;
            }
            else
            {
                return Convert.ToSingle(rs.Fields[strName].Value);
            }
        }
        public static Guid ToGuid(ADODB.Recordset rs, string strName)
        {

            if (ValidRsValue(rs, strName) == false)
            {
                return Guid.Empty;
            }
            else
            {
                return new Guid(rs.Fields[strName].Value.ToString());
            }
        }
        public static decimal ToDecimal(ADODB.Recordset rs, string strName)
        {
            if (ValidRsValue(rs, strName) == false)
            {
                return 0;
            }
            else
            {
                return Convert.ToDecimal(rs.Fields[strName].Value);
            }
        }
        public static double ToDouble(ADODB.Recordset rs, string strName)
        {
            if (ValidRsValue(rs, strName) == false)
            {
                return 0;
            }
            else
            {
                return Convert.ToDouble(rs.Fields[strName].Value);
            }
        }
        public static DateTime ToDateTime(ADODB.Recordset rs, string strName)
        {
            if (ValidRsValue(rs, strName) == false)
            {
                return DateTime.MinValue;
            }
            else
            {
                return Convert.ToDateTime(rs.Fields[strName].Value);
            }
        }
        public static object ToObject(ADODB.Recordset rs, string strName, PropertyInfo property)
        {
            object obj = null;
            obj = ToString(rs, strName);
            switch (Type.GetTypeCode(property.PropertyType))
            {
                case TypeCode.Byte: obj = ToByte(rs, strName); break;
                case TypeCode.Boolean: obj = ToBoolean(rs, strName); break;
                case TypeCode.Char: obj = ToChar(rs, strName); break;
                case TypeCode.Decimal: obj = ToDecimal(rs, strName); break;
                case TypeCode.Double: obj = ToDouble(rs, strName); break;
                case TypeCode.DateTime: obj = ToDateTime(rs, strName); break;
                case TypeCode.Int16: obj = ToInt16(rs, strName); break;
                case TypeCode.Int32: obj = ToInt32(rs, strName); break;
                case TypeCode.Int64: obj = ToInt64(rs, strName); break;
                case TypeCode.Object: obj = ToGuid(rs, strName); break;
                case TypeCode.Single: obj = ToSingle(rs, strName); break;
                case TypeCode.String: obj = ToString(rs, strName); break;
                default: obj = ToString(rs, strName);
                    break;
            }

            return obj;
        }
        public static string AssignGuidToRs(Guid gValue)
        {
            if (gValue.Equals(Guid.Empty))
            {
                return null;
            }
            else
            {
                return "{" + gValue + "}";
            }

        }
        public static void AssignRsGuid(ADODB.Recordset rs, string strName, Guid gValue)
        {
            if (gValue.Equals(Guid.Empty) == false && FieldExist(rs, strName))
            {
                rs.Fields[strName].Value = RecordsetHelper.AssignGuidToRs(gValue);
            }
        }
        public static void AssignRsString(ADODB.Recordset rs, string strName, string strValue)
        {
            if (string.IsNullOrEmpty(strValue) == false && FieldExist(rs, strName))
            {
                rs.Fields[strName].Value = strValue;
            }
        }
        public static void AssignRsDateTime(ADODB.Recordset rs, string strName, DateTime dtValue)
        {
            if (dtValue.Equals(DateTime.MinValue) == false && FieldExist(rs, strName))
            {
                rs.Fields[strName].Value = dtValue;
            }
        }
        public static void AssignRsDecimal(ADODB.Recordset rs, string strName, decimal dValue)
        {
            if (FieldExist(rs, strName))
            {
                rs.Fields[strName].Value = dValue;
            }
        }
        public static void AssignRsDouble(ADODB.Recordset rs, string strName, double dValue)
        {
            if (FieldExist(rs, strName))
            {
                rs.Fields[strName].Value = dValue;
            }
        }
        public static void AssignRsBoolean(ADODB.Recordset rs, string strName, byte bValue)
        {
            if (FieldExist(rs, strName))
            {
                if (bValue == 0)
                    rs.Fields[strName].Value = false;
                else
                    rs.Fields[strName].Value = true;
            }
        }

        private static bool ValidRsValue(ADODB.Recordset rs, string strName)
        {
            if (rs == null || FieldExist(rs, strName) == false || rs.Fields[strName].Value == null || rs.Fields[strName].Value.Equals(System.DBNull.Value))
            { return false; }
            else
            { return true; }
        }
        private static bool FieldExist(ADODB.Recordset rs, string strName)
        {
            if (rs != null && rs.RecordCount > 0)
            {
                for (int i = 0; i < rs.Fields.Count; i++)
                {
                    if (rs.Fields[i].Name.ToUpper().Equals(strName.ToUpper()))
                    {
                        return true;
                    }
                }
            }
            return false;
        }

        public static void ClearRecordset(ref ADODB.Recordset rs)
        {
            ClearRecordset(ref rs, true);
        }
        public static void ClearRecordset(ref ADODB.Recordset rs, bool close)
        {
            if (rs != null)
            {
                if (rs.State == 1 && close == true)
                {
                    rs.Close();
                }
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(rs);
                rs = null;
            }
        }
        public static DataSet RecordsetToDataset(ADODB.Recordset rs, string tableName)
        {
            DataSet ds = new DataSet();
            DataTable dt = new DataTable(tableName);

            try
            {
                foreach (ADODB.Field fld in rs.Fields)
                {
                    Type ttype = fld.GetType();
                    DataColumn dc_column = new DataColumn("column");

                    switch (fld.Type)
                    {
                        case ADODB.DataTypeEnum.adNumeric:
                            dc_column.DataType = System.Type.GetType("System.Double");
                            break;
                        case ADODB.DataTypeEnum.adDouble:
                            dc_column.DataType = System.Type.GetType("System.Double");
                            break;
                        case ADODB.DataTypeEnum.adCurrency:
                            dc_column.DataType = System.Type.GetType("System.Decimal");
                            break;
                        case ADODB.DataTypeEnum.adChar:
                            dc_column.DataType = System.Type.GetType("System.Char");
                            break;
                        case ADODB.DataTypeEnum.adWChar:
                            dc_column.DataType = System.Type.GetType("System.String");
                            break;
                        case ADODB.DataTypeEnum.adVarChar:
                            dc_column.DataType = System.Type.GetType("System.String");
                            break;
                        case ADODB.DataTypeEnum.adLongVarChar:
                            dc_column.DataType = System.Type.GetType("System.String");
                            break;
                        case ADODB.DataTypeEnum.adVarWChar:
                            dc_column.DataType = System.Type.GetType("System.String");
                            break;
                        case ADODB.DataTypeEnum.adLongVarWChar:
                            dc_column.DataType = System.Type.GetType("System.String");
                            break;
                        case ADODB.DataTypeEnum.adBigInt:
                            dc_column.DataType = System.Type.GetType("System.Int64");
                            break;
                        case ADODB.DataTypeEnum.adInteger:
                            dc_column.DataType = System.Type.GetType("System.Int32");
                            break;
                        case ADODB.DataTypeEnum.adSmallInt:
                            dc_column.DataType = System.Type.GetType("System.Int16");
                            break;
                        case ADODB.DataTypeEnum.adUnsignedTinyInt:
                            dc_column.DataType = System.Type.GetType("System.Byte");
                            break;
                        case ADODB.DataTypeEnum.adTinyInt:
                            dc_column.DataType = System.Type.GetType("System.SByte");
                            break;
                        case ADODB.DataTypeEnum.adDBTimeStamp:
                            dc_column.DataType = System.Type.GetType("System.DateTime");
                            break;
                        case ADODB.DataTypeEnum.adDBDate:
                            dc_column.DataType = System.Type.GetType("System.DateTime");
                            break;
                        case ADODB.DataTypeEnum.adDate:
                            dc_column.DataType = System.Type.GetType("System.DateTime");
                            break;
                        case ADODB.DataTypeEnum.adGUID:
                            dc_column.DataType = System.Type.GetType("System.Guid");
                            break;
                        case ADODB.DataTypeEnum.adBoolean:
                            dc_column.DataType = System.Type.GetType("System.Boolean");
                            break;
                        case ADODB.DataTypeEnum.adSingle:
                            dc_column.DataType = System.Type.GetType("System.Single");
                            break;
                        default:
                            System.Diagnostics.Debug.Assert(false);
                            break;
                    }

                    dt.Columns.Add(new DataColumn(fld.Name, dc_column.DataType));

                    ttype = null;
                    if (dc_column != null)
                    {
                        dc_column.Dispose();
                        dc_column = null;
                    }
                }

                object[] vals = new object[rs.Fields.Count];

                rs.MoveFirst();
                while (!rs.EOF)
                {
                    int nFldCnt = 0;

                    foreach (ADODB.Field fld in rs.Fields)
                    {
                        vals[nFldCnt] = fld.Value;
                        nFldCnt += 1;
                    }

                    dt.Rows.Add(vals);

                    rs.MoveNext();
                }
                ds.Tables.Add(dt);
            }
            catch (Exception exp)
            {
                Console.WriteLine(exp);
            }
            return ds;

        }


        public static Recordset FabricateVoucherRecordset()
        {
            Recordset rs = new Recordset();
            rs.Fields.Append("voucher_id", DataTypeEnum.adGUID, 0, ADODB.FieldAttributeEnum.adFldMayBeNull, null);
            rs.Fields.Append("payment_total", DataTypeEnum.adDouble, 0, FieldAttributeEnum.adFldMayBeNull, null);

            rs.Fields.Append("create_by", DataTypeEnum.adGUID, 0, FieldAttributeEnum.adFldIsNullable, null);
            rs.Fields.Append("create_date_time", ADODB.DataTypeEnum.adDate, 0, ADODB.FieldAttributeEnum.adFldMayBeNull, null);
            rs.Fields.Append("update_by", DataTypeEnum.adGUID, 0, FieldAttributeEnum.adFldIsNullable, null);
            rs.Fields.Append("update_date_time", ADODB.DataTypeEnum.adDate, 0, ADODB.FieldAttributeEnum.adFldMayBeNull, null);

            rs.Fields.Append("agency_code", ADODB.DataTypeEnum.adVarChar, 30, ADODB.FieldAttributeEnum.adFldMayBeNull, null);
            rs.Fields.Append("expiry_date_time", ADODB.DataTypeEnum.adDate, 0, ADODB.FieldAttributeEnum.adFldMayBeNull, null);

            rs.Fields.Append("percentage_flag", ADODB.DataTypeEnum.adBoolean, 0, ADODB.FieldAttributeEnum.adFldMayBeNull, null);

            rs.Fields.Append("refundable_flag", ADODB.DataTypeEnum.adBoolean, 0, ADODB.FieldAttributeEnum.adFldMayBeNull, null);

            rs.Fields.Append("single_flight_flag", ADODB.DataTypeEnum.adBoolean, 0, ADODB.FieldAttributeEnum.adFldMayBeNull, null);
            rs.Fields.Append("single_passenger_flag", ADODB.DataTypeEnum.adBoolean, 0, ADODB.FieldAttributeEnum.adFldMayBeNull, null);
            rs.Fields.Append("recipient_name", ADODB.DataTypeEnum.adVarChar, 50, ADODB.FieldAttributeEnum.adFldMayBeNull, null);

            rs.Fields.Append("voucher_number", ADODB.DataTypeEnum.adVarChar, 10, ADODB.FieldAttributeEnum.adFldMayBeNull, null);
            rs.Fields.Append("voucher_password", ADODB.DataTypeEnum.adVarChar, 4, ADODB.FieldAttributeEnum.adFldMayBeNull, null);

            rs.Fields.Append("voucher_status_rcd", ADODB.DataTypeEnum.adVarChar, 4, ADODB.FieldAttributeEnum.adFldMayBeNull, null);
            rs.Fields.Append("voucher_value", DataTypeEnum.adDouble, 0, FieldAttributeEnum.adFldMayBeNull, null);

            //Open Recordset
            rs.CursorLocation = CursorLocationEnum.adUseClient;
            rs.Open(System.Reflection.Missing.Value,
                    System.Reflection.Missing.Value,
                    CursorTypeEnum.adOpenStatic,
                    LockTypeEnum.adLockOptimistic, 0);

            return rs;
        }


        public static Recordset FabricatePaymentAllocationRecordset()
        {
            Recordset rs = new Recordset();
            rs.Fields.Append("booking_payment_id", DataTypeEnum.adGUID, 0, ADODB.FieldAttributeEnum.adFldMayBeNull, null);
            rs.Fields.Append("passenger_id", DataTypeEnum.adGUID, 0, FieldAttributeEnum.adFldMayBeNull, null);
            rs.Fields.Append("booking_id", DataTypeEnum.adGUID, 0, FieldAttributeEnum.adFldMayBeNull, null);
            rs.Fields.Append("booking_segment_id", DataTypeEnum.adGUID, 0, FieldAttributeEnum.adFldMayBeNull, null);
            rs.Fields.Append("voucher_payment_id", DataTypeEnum.adGUID, 0, FieldAttributeEnum.adFldMayBeNull, null);
            rs.Fields.Append("currency_rcd", DataTypeEnum.adVarChar, 20, FieldAttributeEnum.adFldMayBeNull, null);
            rs.Fields.Append("form_of_payment_rcd", DataTypeEnum.adVarChar, 20, FieldAttributeEnum.adFldMayBeNull, null);
            rs.Fields.Append("booking_fee_id", DataTypeEnum.adGUID, 0, FieldAttributeEnum.adFldMayBeNull, null);
            rs.Fields.Append("voucher_id", DataTypeEnum.adGUID, 0, FieldAttributeEnum.adFldMayBeNull, null);
            rs.Fields.Append("fee_id", DataTypeEnum.adGUID, 0, FieldAttributeEnum.adFldMayBeNull, null);
            rs.Fields.Append("user_id", DataTypeEnum.adGUID, 0, FieldAttributeEnum.adFldMayBeNull, null);
            rs.Fields.Append("sales_amount", DataTypeEnum.adDouble, 0, FieldAttributeEnum.adFldMayBeNull, null);
            rs.Fields.Append("payment_amount", DataTypeEnum.adDouble, 0, FieldAttributeEnum.adFldMayBeNull, null);
            rs.Fields.Append("account_amount", DataTypeEnum.adDouble, 0, FieldAttributeEnum.adFldMayBeNull, null);

            rs.Fields.Append("passenger_segment_service_id", DataTypeEnum.adGUID, 0, FieldAttributeEnum.adFldIsNullable, null);
            rs.Fields.Append("od_origin_rcd", DataTypeEnum.adVarChar, 5, FieldAttributeEnum.adFldIsNullable, null);
            rs.Fields.Append("od_destination_rcd", DataTypeEnum.adVarChar, 5, FieldAttributeEnum.adFldIsNullable, null);
            rs.Fields.Append("charge_amount", DataTypeEnum.adDouble, 0, FieldAttributeEnum.adFldIsNullable, null);
            rs.Fields.Append("charge_amount_incl", DataTypeEnum.adDouble, 0, FieldAttributeEnum.adFldIsNullable, null);
            rs.Fields.Append("charge_currency_rcd", DataTypeEnum.adVarChar, 10, FieldAttributeEnum.adFldIsNullable, null);
            rs.Fields.Append("fee_category_rcd", DataTypeEnum.adVarChar, 5, FieldAttributeEnum.adFldIsNullable, null);
            rs.Fields.Append("external_reference", DataTypeEnum.adVarChar, 50, FieldAttributeEnum.adFldIsNullable, null);
            rs.Fields.Append("vendor_rcd", DataTypeEnum.adVarChar, 50, FieldAttributeEnum.adFldIsNullable, null);
            rs.Fields.Append("weight_lbs", DataTypeEnum.adNumeric, 0, FieldAttributeEnum.adFldIsNullable, null);
            rs.Fields.Append("weight_kgs", DataTypeEnum.adNumeric, 0, FieldAttributeEnum.adFldIsNullable, null);
            rs.Fields.Append("units", DataTypeEnum.adNumeric, 0, FieldAttributeEnum.adFldIsNullable, null);
            rs.Fields.Append("create_by", DataTypeEnum.adGUID, 0, FieldAttributeEnum.adFldIsNullable, null);

            //Open Recordset
            rs.CursorLocation = CursorLocationEnum.adUseClient;
            rs.Open(System.Reflection.Missing.Value,
                    System.Reflection.Missing.Value,
                    CursorTypeEnum.adOpenStatic,
                    LockTypeEnum.adLockOptimistic, 0);

            return rs;
        }

        public static Recordset FabricateFlightRecordset()
        {
            Recordset rs = new Recordset();

            rs.Fields.Append("flight_id", ADODB.DataTypeEnum.adGUID, 0, ADODB.FieldAttributeEnum.adFldMayBeNull, null);
            rs.Fields.Append("fare_id", ADODB.DataTypeEnum.adGUID, 0, ADODB.FieldAttributeEnum.adFldMayBeNull, null);
            rs.Fields.Append("booking_segment_id", ADODB.DataTypeEnum.adGUID, 0, ADODB.FieldAttributeEnum.adFldMayBeNull, null);
            rs.Fields.Append("origin_rcd", ADODB.DataTypeEnum.adVarChar, 5, ADODB.FieldAttributeEnum.adFldMayBeNull, null);
            rs.Fields.Append("destination_rcd", ADODB.DataTypeEnum.adVarChar, 5, ADODB.FieldAttributeEnum.adFldMayBeNull, null);
            rs.Fields.Append("airline_rcd", ADODB.DataTypeEnum.adVarChar, 3, ADODB.FieldAttributeEnum.adFldMayBeNull, null);
            rs.Fields.Append("flight_number", ADODB.DataTypeEnum.adVarChar, 5, ADODB.FieldAttributeEnum.adFldMayBeNull, null);
            rs.Fields.Append("boarding_class_rcd", ADODB.DataTypeEnum.adVarChar, 2, ADODB.FieldAttributeEnum.adFldMayBeNull, null);
            rs.Fields.Append("booking_class_rcd", ADODB.DataTypeEnum.adVarChar, 2, ADODB.FieldAttributeEnum.adFldMayBeNull, null);
            rs.Fields.Append("departure_date", ADODB.DataTypeEnum.adDate, 0, ADODB.FieldAttributeEnum.adFldMayBeNull, null);
            rs.Fields.Append("fare_reduction", ADODB.DataTypeEnum.adInteger, 0, ADODB.FieldAttributeEnum.adFldMayBeNull, null);
            rs.Fields.Append("number_of_units", ADODB.DataTypeEnum.adInteger, 0, ADODB.FieldAttributeEnum.adFldMayBeNull, null);
            rs.Fields.Append("waitlist_flag", ADODB.DataTypeEnum.adBoolean, 0, ADODB.FieldAttributeEnum.adFldMayBeNull, null);
            rs.Fields.Append("refundable_flag", ADODB.DataTypeEnum.adBoolean, 0, ADODB.FieldAttributeEnum.adFldMayBeNull, null);
            rs.Fields.Append("non_revenue_flag", ADODB.DataTypeEnum.adBoolean, 0, ADODB.FieldAttributeEnum.adFldMayBeNull, null);
            rs.Fields.Append("eticket_flag", ADODB.DataTypeEnum.adBoolean, 0, ADODB.FieldAttributeEnum.adFldMayBeNull, null);
            rs.Fields.Append("group_flag", ADODB.DataTypeEnum.adBoolean, 0, ADODB.FieldAttributeEnum.adFldMayBeNull, null);
            rs.Fields.Append("overbook_flag", ADODB.DataTypeEnum.adBoolean, 0, ADODB.FieldAttributeEnum.adFldMayBeNull, null);
            rs.Fields.Append("exchanged_segment_id", ADODB.DataTypeEnum.adGUID, 0, ADODB.FieldAttributeEnum.adFldMayBeNull, null);
            rs.Fields.Append("flight_connection_id", ADODB.DataTypeEnum.adGUID, 0, ADODB.FieldAttributeEnum.adFldMayBeNull, null);
            rs.Fields.Append("od_origin_rcd", ADODB.DataTypeEnum.adVarChar, 5, ADODB.FieldAttributeEnum.adFldMayBeNull, null);
            rs.Fields.Append("od_destination_rcd", ADODB.DataTypeEnum.adVarChar, 5, ADODB.FieldAttributeEnum.adFldMayBeNull, null);

            rs.Fields.Append("segment_status_rcd", ADODB.DataTypeEnum.adVarChar, 3, ADODB.FieldAttributeEnum.adFldMayBeNull, null);
            rs.Fields.Append("exclude_quote_flag", ADODB.DataTypeEnum.adUnsignedTinyInt, 0, ADODB.FieldAttributeEnum.adFldMayBeNull, null);
            rs.Fields.Append("advanced_purchase_flag", ADODB.DataTypeEnum.adBoolean, 0, ADODB.FieldAttributeEnum.adFldMayBeNull, null);
            rs.Fields.Append("create_by", ADODB.DataTypeEnum.adGUID, 0, ADODB.FieldAttributeEnum.adFldMayBeNull, null);
            rs.Fields.Append("create_date_time", ADODB.DataTypeEnum.adDBTimeStamp, 0, ADODB.FieldAttributeEnum.adFldMayBeNull, null);
            rs.Fields.Append("update_by", ADODB.DataTypeEnum.adGUID, 0, ADODB.FieldAttributeEnum.adFldMayBeNull, null);
            rs.Fields.Append("update_date_time", ADODB.DataTypeEnum.adDBTimeStamp, 0, ADODB.FieldAttributeEnum.adFldMayBeNull, null);

            rs.Fields.Append("transit_airport_rcd", ADODB.DataTypeEnum.adVarChar, 3, ADODB.FieldAttributeEnum.adFldMayBeNull, null);
            rs.Fields.Append("transit_boarding_class_rcd", ADODB.DataTypeEnum.adVarChar, 2, ADODB.FieldAttributeEnum.adFldMayBeNull, null);
            rs.Fields.Append("transit_booking_class_rcd", ADODB.DataTypeEnum.adVarChar, 2, ADODB.FieldAttributeEnum.adFldMayBeNull, null);
            rs.Fields.Append("transit_flight_id", ADODB.DataTypeEnum.adGUID, 0, ADODB.FieldAttributeEnum.adFldMayBeNull, null);
            rs.Fields.Append("transit_fare_id", ADODB.DataTypeEnum.adGUID, 0, ADODB.FieldAttributeEnum.adFldMayBeNull, null);
            rs.Fields.Append("transit_departure_date", ADODB.DataTypeEnum.adDate, 0, ADODB.FieldAttributeEnum.adFldMayBeNull, null);
            rs.Fields.Append("transit_points", ADODB.DataTypeEnum.adVarChar, 5, ADODB.FieldAttributeEnum.adFldMayBeNull, null);
            rs.Fields.Append("priority_rcd", ADODB.DataTypeEnum.adVarChar, 20, ADODB.FieldAttributeEnum.adFldIsNullable);
            //Open Recordset
            rs.CursorLocation = ADODB.CursorLocationEnum.adUseClient;
            rs.Open(System.Reflection.Missing.Value,
                    System.Reflection.Missing.Value,
                    CursorTypeEnum.adOpenStatic,
                    LockTypeEnum.adLockOptimistic, 0);

            return rs;
        }

        public static void DeleteFees(ref ADODB.Recordset rsFees)
        {
            // clear rs  before  get  new fee
            if (rsFees != null && rsFees.RecordCount > 0)
            {
                rsFees.MoveFirst();
                while (!rsFees.EOF)
                {
                    if (rsFees.Status != 0)
                    {
                        rsFees.Delete();
                    }
                    rsFees.MoveNext();
                }
            }
        }

        public static void DeleteHeader(ref ADODB.Recordset rsPassenger)
        {
            // clear rs  before  get  new fee
            if (rsPassenger != null && rsPassenger.RecordCount > 0)
            {
                rsPassenger.MoveFirst();
                while (!rsPassenger.EOF)
                {
                    rsPassenger.Delete();
                    rsPassenger.MoveNext();
                }
            }
        }


    }
}
