using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ServiceModel;

namespace Avantik.Web.Service.Message.OrderBooking
{
    [MessageContract]
    public class FlightSegment
    {
        [MessageBodyMember]
        public Guid booking_segment_id { get; set; }
        [MessageBodyMember]
        public Guid flight_connection_id { get; set; }
        [MessageBodyMember]
        public Guid flight_id { get; set; }
        [MessageBodyMember]
        public DateTime departure_date { get; set; }
        [MessageBodyMember]
        public int departure_time { get; set; }
        [MessageBodyMember]
        public DateTime arrival_date { get; set; }
        [MessageBodyMember]
        public int arrival_time { get; set; }
        [MessageBodyMember]
        public string airline_rcd { get; set; }
        [MessageBodyMember]
        public string flight_number { get; set; }
        [MessageBodyMember]
        public string origin_rcd { get; set; }
        [MessageBodyMember]
        public string destination_rcd { get; set; }
        [MessageBodyMember]
        public string od_origin_rcd { get; set; }
        [MessageBodyMember]
        public string od_destination_rcd { get; set; }
        [MessageBodyMember]
        public string booking_class_rcd { get; set; }
        [MessageBodyMember]
        public string boarding_class_rcd { get; set; }
        [MessageBodyMember]
        public string segment_status_rcd { get; set; }
        [MessageBodyMember]
        public int number_of_units { get; set; }

        [MessageBodyMember]
        public Guid exchanged_segment_id { get; set; }

        [MessageBodyMember]
        public Guid fare_id { get; set; }

    }
}
