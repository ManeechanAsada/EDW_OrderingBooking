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
    public class BookingPaymentRequest
    {
        [MessageHeader]
        public string Token { get; set; }
        [MessageBodyMember]
        public IList<Mapping> Mappings { get; set; }
        [MessageBodyMember]
        public IList<Fee> Fees { get; set; }
        [MessageBodyMember]
        public IList<Payment> Payments { get; set; }
    }
}
