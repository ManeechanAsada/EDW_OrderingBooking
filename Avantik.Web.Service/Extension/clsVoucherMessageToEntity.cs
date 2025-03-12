using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Avantik.Web.Service.Entity;

namespace Avantik.Web.Service.Extension
{
    public static class VoucherMessageToEntity
    {
        public static Voucher ToVoucherEntity(this Message.Voucher vc)
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

                voucher.ValidateVoucherDuplicate(null);
            }
            return voucher;
        }

        public static IList<Entity.Voucher> ToVoucherEntity(this IList<Message.Voucher> objMessageVouchers)
        {
            List<Entity.Voucher> objVoucherEntities = null;
            if (objMessageVouchers != null)
            {
                objVoucherEntities = new List<Entity.Voucher>();
                for (int i = 0; i < objMessageVouchers.Count; i++)
                {
                    objVoucherEntities.Add(objMessageVouchers[i].ToVoucherEntity());
                }
            }
            return objVoucherEntities;
        }
    }
}
