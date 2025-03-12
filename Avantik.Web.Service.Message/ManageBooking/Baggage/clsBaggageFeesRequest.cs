using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.Text;

namespace Avantik.Web.Service.Message.ManageBooking
{
    [MessageContract]
    public class BaggageFeesRequest
    {
        [MessageHeader]
        public string Token { get; set; }

        [MessageBodyMember]
        public string BookingId { set; get; }
        [MessageBodyMember]
        public IList<Baggage> Baggage { get; set; }
    }
}
