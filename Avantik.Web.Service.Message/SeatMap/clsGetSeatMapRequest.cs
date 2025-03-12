using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using System.ServiceModel;
using System.Runtime.Serialization;
using Avantik.Web.Service;

namespace Avantik.Web.Service.Message.SeatMap
{
    [MessageContract]
    public class GetSeatMapRequest
    {
        [MessageHeader]
        public string Token { get; set; }

        [MessageBodyMember]
        public string OriginRcd { get; set; }
        [MessageBodyMember]
        public string DestinationRcd { get; set; }
        [MessageBodyMember]
        public string FlightId { get; set; }
        [MessageBodyMember]
        public string BoardingClass { get; set; }
        [MessageBodyMember]
        public string BookingClass { get; set; }
    }
}
