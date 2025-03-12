using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Avantik.Web.Service.Entity.Agency;

namespace Avantik.Web.Service.Entity.Booking
{
    public class Payment
    {
        #region Property
        public Guid BookingPaymentId { get; set; }
        public Guid BookingSegmentId { get; set; }
        public Guid BookingId { get; set; }
        public Guid VoucherPaymentId { get; set; }
        public string FormOfPaymentRcd { get; set; }
        public string CurrencyRcd { get; set; }
        public string ReceiveCurrencyRcd { get; set; }
        public string AgencyPaymentTypeRcd { get; set; }
        public string AgencyCode { get; set; }
        public string DebitAgencyCode { get; set; }
        public decimal PaymentAmount { get; set; }
        public decimal ReceivePaymentAmount { get; set; }
        public decimal AcctPaymentAmount { get; set; }
        public Guid PaymentBy { get; set; }
        public DateTime PaymentDateTime { get; set; }
        public DateTime PaymentDueDateTime { get; set; }
        public decimal DocumentAmount { get; set; }
        public Guid VoidBy { get; set; }
        public int ExpiryMonth { get; set; }
        public int ExpiryYear { get; set; }
        public DateTime VoidDateTime { get; set; }
        public string RecordLocator { get; set; }
        public string CvvCode { get; set; }
        public string NameOnCard { get; set; }
        public string DocumentNumber { get; set; }
        public string DocumentPassword { get; set; }
        public string FormOfPaymentSubtypeRcd { get; set; }
        public string City { get; set; }
        public string State { get; set; }
        public string Street { get; set; }
        public string PoBox { get; set; }
        public string AddressLine1 { get; set; }
        public string AddressLine2 { get; set; }
        public string District { get; set; }
        public string Province { get; set; }
        public string ZipCode { get; set; }
        public string CountryRcd { get; set; }
        public Guid CreateBy { get; set; }
        public DateTime CreateDateTime { get; set; }
        public Guid UpdateBy { get; set; }
        public DateTime UpdateDateTime { get; set; }
        public string PosIndicator { get; set; }
        public int IssueMonth { get; set; }
        public int IssueYear { get; set; }
        public string IssueNumber { get; set; }
        public Guid ClientProfileId { get; set; }
        public string TransactionReference { get; set; }
        public string ApprovalCode { get; set; }
        public string BankName { get; set; }
        public string BankCode { get; set; }
        public string BankIban { get; set; }
        public string IpAddress { get; set; }
        public string PaymentReference { get; set; }
        public decimal AllocatedAmount { get; set; }
        public string PaymentTypeRcd { get; set; }
        public decimal RefundedAmount { get; set; }

        public string PaymentRemark { get; set; }
        public string PaymentNumber { get; set; }
        public string PaymentNumberPassword { get; set; }

        #endregion

        #region Method
        public bool ValidateCreditAgency(Agent agent)
        {
            bool bResult = true;
            decimal agentBalance = 0;
            try
            {
                if (!string.IsNullOrEmpty(this.FormOfPaymentRcd))
                {
                    if (this.FormOfPaymentRcd != "CRAGT" && this.FormOfPaymentRcd != "INV")
                    {
                        bResult = false;
                    }
                }

                if (agent != null)
                {
                    if (agent.CurrencyRcd != this.CurrencyRcd)
                    {
                        bResult = false;
                    }
                    
                    agentBalance = ((agent.AgencyAccount - agent.BookingPayment) - this.PaymentAmount);
                    if (agentBalance <= 0)
                    {
                        bResult = false;
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return bResult;
        }
        public static int ComparePaymentAmount(Payment value1, Payment value2)
        {
            return value1.PaymentAmount.CompareTo(value2.PaymentAmount);
        }
        #endregion

    }
}
