using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.Serialization;
using System.ServiceModel;

namespace Avantik.Web.Service.Message.Fee
{
    public class SegmentService
    {
        [MessageBodyMember]
        public Guid FlightConnectionId { get; set; }
        [MessageBodyMember]
        public string SpecialServiceRcd { get; set; }
        [MessageBodyMember]
        public string OriginRcd { get; set; }
        [MessageBodyMember]
        public string DestinationRcd { get; set; }
        [MessageBodyMember]
        public string OdOriginRcd { get; set; }
        [MessageBodyMember]
        public string OdDestinationRcd { get; set; }
        [MessageBodyMember]
        public string BookingClassRcd { get; set; }
        [MessageBodyMember]
        public string FareCode { get; set; }
        [MessageBodyMember]
        public string AirlineRcd { get; set; }
        [MessageBodyMember]
        public string FlightNumber { get; set; }
        [DataMember(EmitDefaultValue = false)]
        public DateTime DepartureDate { get; set; }
    }
}
