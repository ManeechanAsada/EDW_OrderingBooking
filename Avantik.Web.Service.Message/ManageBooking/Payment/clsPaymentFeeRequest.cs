using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.Text;

namespace Avantik.Web.Service.Message
{
    [MessageContract]
    public class PaymentFeeRequest
    {
        [MessageHeader]
        public string Token { get; set; }

        [MessageBodyMember]
        public Message.ManageBooking.PaymentFee PaymentFee { get; set; }
    }
}
