using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.ServiceModel;

namespace Avantik.Web.Service.Message.OrderBooking
{
    [MessageContract]
    public class Service
    {
        [MessageBodyMember]
        public Guid passenger_id { get; set; }
        [MessageBodyMember]
        public Guid booking_segment_id { get; set; }
        [MessageBodyMember]
        public short number_of_units { get; set; }
        [MessageBodyMember]
        public string special_service_rcd { get; set; }
        [MessageBodyMember]
        public string service_text { get; set; }
        [MessageBodyMember]
        public string unit_rcd { get; set; }
    }
}
