using Avantik.Web.Service.COMHelper;
using Avantik.Web.Service.Entity.Booking;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Avantik.Web.Service.Model.COM.Extension
{
    public static class RecordsetObjectBaggage
    {
        public static void FillBaggage(this IList<Entity.Booking.Fee> fees, ref ADODB.Recordset rs)
        {
            if (rs != null && rs.RecordCount > 0)
            {

                Entity.Booking.Fee bkfee = null;

                try
                {
                    rs.MoveFirst();
                    while (!rs.EOF)
                    {
                        bkfee = new Fee();
                        bkfee.BaggageFeeOptionId = RecordsetHelper.ToGuid(rs, "baggage_fee_option_id");
                        bkfee.PassengerId = RecordsetHelper.ToGuid(rs, "passenger_id");
                        bkfee.BookingSegmentId = RecordsetHelper.ToGuid(rs, "booking_segment_id");
                        bkfee.FeeId = RecordsetHelper.ToGuid(rs, "fee_id");
                        bkfee.FeeRcd = RecordsetHelper.ToString(rs, "fee_rcd");
                        bkfee.FeeCategoryRcd = RecordsetHelper.ToString(rs, "fee_category_rcd");
                        bkfee.CurrencyRcd = RecordsetHelper.ToString(rs, "currency_rcd");
                        bkfee.DisplayName = RecordsetHelper.ToString(rs, "display_name");
                        bkfee.NumberOfUnits = RecordsetHelper.ToDecimal(rs, "number_of_units");
                        bkfee.FeeAmount = RecordsetHelper.ToDecimal(rs, "fee_amount");
                        bkfee.AcctFeeAmountIncl = RecordsetHelper.ToDecimal(rs, "fee_amount_incl");
                        bkfee.TotalAmount = RecordsetHelper.ToDecimal(rs, "total_amount");
                        bkfee.TotalAmountIncl = RecordsetHelper.ToDecimal(rs, "total_amount_incl");
                        bkfee.VatPercentage = RecordsetHelper.ToDecimal(rs, "vat_percentage");

                        fees.Add(bkfee);
                        rs.MoveNext();
                    }

                }
                catch
                {
                    throw;
                }
            }

        }
    }
}
