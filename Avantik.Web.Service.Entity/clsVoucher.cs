using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Avantik.Web.Service.Entity;
using Avantik.Web.Service.Entity.Booking;

namespace Avantik.Web.Service.Entity
{
    public class Voucher : VoucherTemplate
    {
        #region Property
        public string AgencyCode { get; set; }
        public Guid CreateBy { get; set; }
        public DateTime CreateDateTime { get; set; }
        public DateTime ExpiryDateTime { get; set; }
        public decimal PaymentTotal { get; set; }
        public byte PercentageFlag { get; set; }
        public string RecipientName { get; set; }
        public byte RefundableFlag { get; set; }
        public byte SingleFlightFlag { get; set; }
        public byte SinglePassengerFlag { get; set; }
        public Guid UpdateBy { get; set; }
        public DateTime UpdateDateTime { get; set; }
        public string VoucherId { get; set; }
        public string VoucherNumber { get; set; }
        public string VoucherPassword { get; set; }
        public string VoucherStatusRcd { get; set; }
        #endregion

        #region Method
        public bool ValidateVoucherDuplicate(IList<Payment> payments)
        {
            bool bResult = true;
            try
            {
                if (payments != null && payments.Count > 0)
                {
                    if (payments.Where(b=>b.FormOfPaymentRcd == "VOUCHER").GroupBy(a=>a.DocumentNumber).Any(c => c.Count() > 1))
                    {
                        bResult = false;
                    }
                }
                else
                {
                    bResult = false;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return bResult;
        }
        public decimal GetVouchersPaymentBalance(IList<Payment> payments)
        {
            decimal dBalance = 0;
            try
            {
                if (payments != null && payments.Count > 0)
                {
                    dBalance = payments.Where(a => a.FormOfPaymentRcd == "VOUCHER").Sum(b => b.DocumentAmount);
                }
            }
            catch(Exception ex)
            {
                throw ex;
            }
            return dBalance;
        }
        public bool ValidateVoucherPayment(IList<Payment> payments)
        {
            bool bResult = false;
            decimal dVoucherAmount = 0;
            decimal dPaymentAmount = 0;
            try
            {
                if (payments != null && payments.Count > 0)
                {
                    dVoucherAmount = GetVouchersPaymentBalance(payments);
                    dPaymentAmount = payments.Sum(b => b.PaymentAmount);

                    if (dVoucherAmount - dPaymentAmount >= 0)
                    {
                        bResult = true;
                    }
                }
            }
            catch(Exception ex)
            {
                throw ex;
            }
            return bResult;
        }

        public bool ValidateVoucherEnough(decimal totalPeyment)
        {
            bool bResult = false;
            decimal dVoucherValue = 0;
            try
            {
                dVoucherValue = this.VoucherValue - this.PaymentTotal;

                if (dVoucherValue - totalPeyment >= 0)
                {
                    bResult = true;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return bResult;
        }
        #endregion
    }
}
