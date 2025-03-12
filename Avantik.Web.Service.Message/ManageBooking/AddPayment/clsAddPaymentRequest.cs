using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using System.ServiceModel;
using System.Runtime.Serialization;
using Avantik.Web.Service;
using Avantik.Web.Service.Message.Fee;

namespace Avantik.Web.Service.Message.ManageBooking
{
    [MessageContract]
    public class AddPaymentRequest
    {
        [MessageHeader]
        public string Token { get; set; }
        [MessageBodyMember]
        public string BookingId { set; get; }
        [MessageBodyMember]
        public IList<Payment> Payments { get; set; }
    }
}
