using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.ServiceModel;
using System.Text;
using Avantik.Web.Service.Message.ManageBooking;

namespace Avantik.Web.Service.Message.ManageBooking
{
    [MessageContract]
    public class SegmentCancelRequest
    {
        [MessageHeader]
        public string Token { get; set; }
        [MessageBodyMember]
        public string BookingId { set; get; }
        [MessageBodyMember]
        public string SegmentId { set; get; }
    }
}
