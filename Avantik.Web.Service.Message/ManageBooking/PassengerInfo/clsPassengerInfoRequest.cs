using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.Text;

namespace Avantik.Web.Service.Message
{
    [MessageContract]
    public class PassengerInfoRequest
    {
        [MessageHeader]
        public string Token { get; set; }

        [MessageBodyMember]
        public string BookingId { set; get; }
        [MessageBodyMember]
        public IList<PassengerInfo> PassengerInfo { get; set; }
    }
}
