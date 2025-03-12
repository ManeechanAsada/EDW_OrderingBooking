using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.Serialization;
using System.ServiceModel;

namespace Avantik.Web.Service.Message.OrderBooking
{
    [MessageContract]
    public class OrderBookingRequest
    {
        [MessageHeader]
        public string Token { set; get; }

        [MessageBodyMember]
        public string BookingId { set; get; }
        [MessageBodyMember]
        public IList<Message.OrderBooking.FlightSegment> BookingSegments { get; set; }
        [MessageBodyMember]
        public IList<Message.OrderBooking.Mapping> Mappings { get; set; }
        [MessageBodyMember]
        public IList<Message.OrderBooking.Fee> Fees { get; set; }
        [MessageBodyMember]
        public IList<Message.OrderBooking.Service> Services { get; set; }
        [MessageBodyMember]
        public IList<Message.OrderBooking.Tax> Taxes { get; set; }

    }
}
