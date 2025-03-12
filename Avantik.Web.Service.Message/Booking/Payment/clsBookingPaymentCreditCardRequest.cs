using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using System.ServiceModel;
using System.Runtime.Serialization;
using Avantik.Web.Service;

namespace Avantik.Web.Service.Message.Booking
{
    [MessageContract]
    public class BookingPaymentCreditCardRequest 
    {
        [MessageBodyMember]
        public BookingRequest Booking { get; set; }
        [MessageBodyMember]
        public Fee PaymentFee { get; set; }
        [MessageBodyMember]
        public string SecurityToken { get; set; }
        [MessageBodyMember]
        public string AuthenticationToken { get; set; }
        [MessageBodyMember]
        public string CommerceIndicator { get; set; }
        [MessageBodyMember]
        public string BookingReference { get; set; }
        [MessageBodyMember]
        public string RequestSource { get; set; }
    }
}
