using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.Serialization;
using System.ServiceModel;

namespace Avantik.Web.Service.Message.Booking
{
    public class Payment
    {
        [MessageBodyMember]
        public Guid BookingPaymentId { get; set; }
        [MessageBodyMember]
        public Guid BookingSegmentId { get; set; }
        [MessageBodyMember]
        public Guid BookingId { get; set; }
        [MessageBodyMember]
        public Guid VoucherPaymentId { get; set; }
        [MessageBodyMember]
        public string FormOfPaymentRcd { get; set; }
        [MessageBodyMember]
        public string CurrencyRcd { get; set; }
        [MessageBodyMember]
        public string ReceiveCurrencyRcd { get; set; }
        [MessageBodyMember]
        public string AgencyPaymentTypeRcd { get; set; }
        [MessageBodyMember]
        public string AgencyCode { get; set; }
        [MessageBodyMember]
        public string DebitAgencyCode { get; set; }
        [MessageBodyMember]
        public decimal PaymentAmount { get; set; }
        [MessageBodyMember]
        public decimal ReceivePaymentAmount { get; set; }
        [MessageBodyMember]
        public decimal AcctPaymentAmount { get; set; }
        [MessageBodyMember]
        public Guid PaymentBy { get; set; }
        [MessageBodyMember]
        public DateTime PaymentDateTime { get; set; }
        [MessageBodyMember]
        public DateTime PaymentDueDateTime { get; set; }
        [MessageBodyMember]
        public decimal DocumentAmount { get; set; }
        [MessageBodyMember]
        public Guid VoidBy { get; set; }
        [MessageBodyMember]
        public int ExpiryMonth { get; set; }
        [MessageBodyMember]
        public int ExpiryYear { get; set; }
        [MessageBodyMember]
        public DateTime VoidDateTime { get; set; }
        [MessageBodyMember]
        public string RecordLocator { get; set; }
        [MessageBodyMember]
        public string CvvCode { get; set; }
        [MessageBodyMember]
        public string NameOnCard { get; set; }
        [MessageBodyMember]
        public string DocumentNumber { get; set; }
        [MessageBodyMember]
        public string FormOfPaymentSubtypeRcd { get; set; }
        [MessageBodyMember]
        public string City { get; set; }
        [MessageBodyMember]
        public string State { get; set; }
        [MessageBodyMember]
        public string Street { get; set; }
        [MessageBodyMember]
        public string PoBox { get; set; }
        [MessageBodyMember]
        public string AddressLine1 { get; set; }
        [MessageBodyMember]
        public string AddressLine2 { get; set; }
        [MessageBodyMember]
        public string ZipCode { get; set; }
        [MessageBodyMember]
        public string District { get; set; }
        [MessageBodyMember]
        public string Province { get; set; }

        [MessageBodyMember]
        public string CountryRcd { get; set; }
        [MessageBodyMember]
        public Guid CreateBy { get; set; }
        [MessageBodyMember]
        public DateTime CreateDateTime { get; set; }
        [MessageBodyMember]
        public Guid UpdateBy { get; set; }
        [MessageBodyMember]
        public DateTime UpdateDateTime { get; set; }
        [MessageBodyMember]
        public string PosIndicator { get; set; }
        [MessageBodyMember]
        public int IssueMonth { get; set; }
        [MessageBodyMember]
        public int IssueYear { get; set; }
        [MessageBodyMember]
        public string IssueNumber { get; set; }
        [MessageBodyMember]
        public Guid ClientProfileId { get; set; }
        [MessageBodyMember]
        public string TransactionReference { get; set; }
        [MessageBodyMember]
        public string ApprovalCode { get; set; }
        [MessageBodyMember]
        public string BankName { get; set; }
        [MessageBodyMember]
        public string BankCode { get; set; }
        [MessageBodyMember]
        public string BankIban { get; set; }
        [MessageBodyMember]
        public string IpAddress { get; set; }
        [MessageBodyMember]
        public string PaymentReference { get; set; }
        [MessageBodyMember]
        public decimal AllocatedAmount { get; set; }
        [MessageBodyMember]
        public string PaymentTypeRcd { get; set; }
        [MessageBodyMember]
        public decimal RefundedAmount { get; set; }
        [MessageBodyMember]
        public string PaymentNumber { get; set; }
        [MessageBodyMember]
        public string PaymentRemark { get; set; }

    }
}
