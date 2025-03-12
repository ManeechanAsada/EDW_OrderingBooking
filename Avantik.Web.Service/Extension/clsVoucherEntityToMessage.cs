using Avantik.Web.Service.Message;
using Avantik.Web.Service.Message.Booking;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Avantik.Web.Service.Extension
{
    public static class VoucherEntityToMessage
    {
        public static Voucher ToVoucherMessage(this Avantik.Web.Service.Entity.Voucher vc)
        {
            Voucher voucher = null;

            if (vc != null)
            {
                voucher = new Voucher();
                voucher.AgencyCode = vc.AgencyCode;
                voucher.AirlineFlag = vc.AirlineFlag;
                voucher.B2bFlag = vc.B2bFlag;
                voucher.B2cFlag = vc.B2cFlag;
                voucher.B2eFlag = vc.B2eFlag;
                voucher.ChargeAmount = vc.ChargeAmount;
                voucher.CreateBy = vc.CreateBy;
                voucher.CreateDateTime = vc.CreateDateTime;
                voucher.CurrencyRcd = vc.CurrencyRcd;
                voucher.Destinations = vc.Destinations;
                voucher.DiscountPercentage = vc.DiscountPercentage;
                voucher.DisplayName = vc.DisplayName;
                voucher.ExpiryDateTime = vc.ExpiryDateTime;
                voucher.FareOnlyFlag = vc.FareOnlyFlag;
                voucher.FormOfPaymentRcd = vc.FormOfPaymentRcd;
                voucher.FormOfPaymentSubtypeRcd = vc.FormOfPaymentSubtypeRcd;
                voucher.MultiplePaymentFlag = vc.MultiplePaymentFlag;
                voucher.Origins = vc.Origins;
                voucher.OtherFeeFlag = vc.OtherFeeFlag;
                voucher.PassengerSegments = vc.PassengerSegments;
                voucher.PaymentTotal = vc.PaymentTotal;
                voucher.PercentageFlag = vc.PercentageFlag;
                voucher.RecipientName = vc.RecipientName;
                voucher.RecipientOnlyFlag = vc.RecipientOnlyFlag;
                voucher.RefundableFlag = vc.RefundableFlag;
                voucher.SeatFeeFlag = vc.SeatFeeFlag;
                voucher.SingleFlightFlag = vc.SingleFlightFlag;
                voucher.SinglePassengerFlag = vc.SinglePassengerFlag;
                voucher.StatusCode = vc.StatusCode;
                voucher.TicketFlag = vc.TicketFlag;
                voucher.UpdateBy = vc.UpdateBy;
                voucher.UpdateDateTime = vc.UpdateDateTime;
                voucher.ValidDays = vc.ValidDays;
                voucher.ValidForClass = vc.ValidForClass;
                voucher.ValidFromDate = vc.ValidFromDate;
                voucher.ValidToDate = vc.ValidToDate;
                voucher.VoucherId = vc.VoucherId;
                voucher.VoucherNumber = vc.VoucherNumber;
                voucher.VoucherPassword = vc.VoucherPassword;
                voucher.VoucherStatusRcd = vc.VoucherStatusRcd;
                voucher.VoucherTemplateId = vc.VoucherTemplateId;
                voucher.VoucherUseCode = vc.VoucherUseCode;
                voucher.VoucherValue = vc.VoucherValue;            
            }
            return voucher;
        }

        public static IList<Message.Voucher> ToVoucherMessage(this IList<Entity.Voucher> objEntityVouchers)
        {
            List<Message.Voucher> objMessageVouchers = null;
            if (objEntityVouchers != null)
            {
                objMessageVouchers = new List<Message.Voucher>();
                for (int i = 0; i < objEntityVouchers.Count; i++)
                {
                    objMessageVouchers.Add(objEntityVouchers[i].ToVoucherMessage());
                }
            }
            return objMessageVouchers;
        }
    }
}
