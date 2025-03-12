using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Avantik.Web.Service.Entity;
using Avantik.Web.Service.COMHelper;
using Avantik.Web.Service.Exception.Booking;

namespace Avantik.Web.Service.Model.COM.Extension
{
    public static class RecordsetObjectVoucher
    {
        public static Voucher FillVoucher(this Voucher voucher, ADODB.Recordset rs)
        {
            voucher = new Voucher();
            try
            {
                voucher.AgencyCode = RecordsetHelper.ToString(rs, "agency_code");
                voucher.CreateBy = RecordsetHelper.ToGuid(rs, "create_by");
                voucher.CreateDateTime = RecordsetHelper.ToDateTime(rs, "create_date_time");
                voucher.ExpiryDateTime = RecordsetHelper.ToDateTime(rs, "expiry_date_time");
                voucher.PaymentTotal = RecordsetHelper.ToDecimal(rs, "payment_total");
                voucher.PercentageFlag = RecordsetHelper.ToByte(rs, "percentage_flag");
                voucher.RecipientName = RecordsetHelper.ToString(rs, "recipient_name");
                voucher.RefundableFlag = RecordsetHelper.ToByte(rs, "refundable_flag");
                voucher.SingleFlightFlag = RecordsetHelper.ToByte(rs, "single_flight_flag");
                voucher.SinglePassengerFlag = RecordsetHelper.ToByte(rs, "single_passenger_flag");
                voucher.UpdateBy = RecordsetHelper.ToGuid(rs, "update_by");
                voucher.UpdateDateTime = RecordsetHelper.ToDateTime(rs, "update_date_time");
                voucher.VoucherId = RecordsetHelper.ToString(rs, "voucher_id");
                voucher.VoucherNumber = RecordsetHelper.ToString(rs, "voucher_number");
                voucher.VoucherPassword = RecordsetHelper.ToString(rs, "voucher_password");
                voucher.VoucherStatusRcd = RecordsetHelper.ToString(rs, "voucher_status_rcd");
                voucher.VoucherValue = RecordsetHelper.ToDecimal(rs, "voucher_value");
            }
            catch
            {
                throw;
            }
            return voucher;
        }
    }
}
