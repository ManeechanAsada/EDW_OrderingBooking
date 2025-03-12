using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.ServiceModel;
using System.Text;

namespace Avantik.Web.Service.Message.ManageBooking
{
    public class NewSegment
    {
        [MessageBodyMember]
        public string BoardingClassRcd { get; set; }
        [MessageBodyMember]
        public string BookingClassRcd { get; set; }
        [DataMember(EmitDefaultValue = false)]
        public DateTime DepartureDate { get; set; }
        [MessageBodyMember]
        public string OriginRcd { get; set; }
        [MessageBodyMember]
        public string DestinationRcd { get; set; }
        [MessageBodyMember]
        public string OdOriginRcd { get; set; }
        [MessageBodyMember]
        public string OdDestinationRcd { get; set; }
        [MessageBodyMember]
        public string FareId { get; set; }
        [MessageBodyMember]
        public string FlightId { get; set; }
        [MessageBodyMember]
        public string TransitFlightId { get; set; }
        [MessageBodyMember]
        public string TransitAirportRcd { get; set; }
        [MessageBodyMember]
        public string TransitBoardingClassRcd { get; set; }
        [MessageBodyMember]
        public string TransitBookingClassRcd { get; set; }
        [MessageBodyMember]
        public DateTime TransitDepartureDate { get; set; }
        [MessageBodyMember]
        public string TransitFareId { get; set; }
        
    }
}
