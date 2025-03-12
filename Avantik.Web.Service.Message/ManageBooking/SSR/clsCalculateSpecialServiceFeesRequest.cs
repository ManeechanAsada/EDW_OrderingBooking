using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using System.ServiceModel;
using System.Runtime.Serialization;
using Avantik.Web.Service;
using Avantik.Web.Service.Message.Booking;

namespace Avantik.Web.Service.Message
{
    [MessageContract]
    public class CalculateSpecialServiceFeesRequest
    {
        [MessageHeader]
        public string Token { set; get; }
        [MessageBodyMember]
        public string BookingId { set; get; }
        [MessageBodyMember]
        public IList<Message.SSR.Service> Services { set; get; }
    }
}
