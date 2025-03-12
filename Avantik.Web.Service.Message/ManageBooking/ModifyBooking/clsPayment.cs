using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.Serialization;
using System.ServiceModel;

namespace Avantik.Web.Service.Message.ManageBooking
{
    public class Payment
    {
        [MessageBodyMember]
        public string FormOfPaymentRcd { get; set; }
       // [MessageBodyMember]
       // public string CurrencyRcd { get; set; }
       // [MessageBodyMember]
      //  public string ReceiveCurrencyRcd { get; set; }
       // [MessageBodyMember]
      //  public DateTime PaymentDueDateTime { get; set; }
        [MessageBodyMember]
        public int ExpiryMonth { get; set; }
        [MessageBodyMember]
        public int ExpiryYear { get; set; }
       // [MessageBodyMember]
       // public string RecordLocator { get; set; }
        [MessageBodyMember]
        public string CvvCode { get; set; }
        [MessageBodyMember]
        public string NameOnCard { get; set; }
        [MessageBodyMember]
        public string DocumentNumber { get; set; }
        [MessageBodyMember]
        public string DocumentPassword { get; set; }
        [MessageBodyMember]
        public string FormOfPaymentSubtypeRcd { get; set; }
        [MessageBodyMember]
        public string City { get; set; }
        [MessageBodyMember]
        public string State { get; set; }
        [MessageBodyMember]
        public string Street { get; set; }
       // [MessageBodyMember]
      //  public string PoBox { get; set; }
        [MessageBodyMember]
        public string AddressLine1 { get; set; }
     //   [MessageBodyMember]
     //   public string AddressLine2 { get; set; }
        [MessageBodyMember]
        public string ZipCode { get; set; }
     //   [MessageBodyMember]
     //   public string District { get; set; }
       // [MessageBodyMember]
       // public string Province { get; set; }
        [MessageBodyMember]
        public string CountryRcd { get; set; }
      //  [MessageBodyMember]
      //  public string PosIndicator { get; set; }
        [MessageBodyMember]
        public int IssueMonth { get; set; }
        [MessageBodyMember]
        public int IssueYear { get; set; }
        [MessageBodyMember]
        public string IssueNumber { get; set; }
        [MessageBodyMember]
        public string IpAddress { get; set; }
       // [MessageBodyMember]
       // public string PaymentReference { get; set; }
     //   [MessageBodyMember]
      //  public decimal RefundedAmount { get; set; }
      //  [MessageBodyMember]
     //   public string PaymentRemark { get; set; }

    }
}
