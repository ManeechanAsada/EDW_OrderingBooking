using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.Serialization;
using System.ServiceModel;

namespace Avantik.Web.Service.Message.Booking
{
    public class Flight
    {
        [MessageBodyMember]
        public string BoardingClassRcd { get; set; }

        [MessageBodyMember]
        public string BookingClassRcd { get; set; }

        [MessageBodyMember]
        public Guid FlightId { get; set; }

        [MessageBodyMember]
        public string OriginRcd { get; set; }

        [DataMember(EmitDefaultValue = false)]
        public DateTime DepartureDate { get; set; }

        [MessageBodyMember]
        public Guid FairId { get; set; }

        [MessageBodyMember]
        public byte EticketFlag { get; set; }

        [MessageBodyMember]
        public string OdOriginRcd { get; set; }

        [MessageBodyMember]
        public string OdDestinationRcd { get; set; }

        [MessageBodyMember]
        public Guid FlightConnectionId { get; set; }

        [MessageBodyMember]
        public string DestinationRcd { get; set; }

        [MessageBodyMember]
        public string TransitPoints { get; set; }

        [MessageBodyMember]
        public string TransitPointsName { get; set; }

        [MessageBodyMember]
        public string AirlineRcd { get; set; }

        [MessageBodyMember]
        public string FlightNumber { get; set; }

        [MessageBodyMember]
        public int NumberOfUnits { get; set; }
    }
}
