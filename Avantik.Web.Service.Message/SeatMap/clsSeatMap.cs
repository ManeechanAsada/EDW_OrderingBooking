using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.Serialization;
using System.ServiceModel;

namespace Avantik.Web.Service.Message.SeatMap
{
    public class SeatMap
    {
        [MessageBodyMember]
        public Guid FlightId { get; set; }
        [MessageBodyMember]
        public int FreeSeatingFlag { get; set; }
        [MessageBodyMember]
        public string FlightCheckInStatusRcd { get; set; }
        [MessageBodyMember]
        public string OriginRcd { get; set; }
        [MessageBodyMember]
        public string DestinationRcd { get; set; }
        [MessageBodyMember]
        public string AircraftConfigurationCode { get; set; }
        [MessageBodyMember]
        public int NumberOfBays { get; set; }
        [MessageBodyMember]
        public string BoardingClassRcd { get; set; }
        [MessageBodyMember]
        public int NumberOfRows { get; set; }
        [MessageBodyMember]
        public int NumberOfColumns { get; set; }
        [MessageBodyMember]
        public int LayoutRow { get; set; }
        [MessageBodyMember]
        public int LayoutColumn { get; set; }
        [MessageBodyMember]
        public string LocationTypeRcd { get; set; }
        [MessageBodyMember]
        public string FeeRcd { get; set; }
        [MessageBodyMember]
        public string SeatColumn { get; set; }
        [MessageBodyMember]
        public int SeatRow { get; set; }
        [MessageBodyMember]
        public int StretcherFlag { get; set; }
        [MessageBodyMember]
        public int HanddicappedFlag { get; set; }
        [MessageBodyMember]
        public int NoChildFlag { get; set; }
        [MessageBodyMember]
        public int BassinetFlag { get; set; }
        [MessageBodyMember]
        public int NoInfantFlag { get; set; }
        [MessageBodyMember]
        public int InfantFlag { get; set; }
        [MessageBodyMember]
        public int EmergencyExitFlag { get; set; }
        [MessageBodyMember]
        public int UnAccompaniedMinorsFlag { get; set; }
        [MessageBodyMember]
        public int WindowFlag { get; set; }
        [MessageBodyMember]
        public int AisleFlag { get; set; }
        [MessageBodyMember]
        public int BlockB2cFlag { get; set; }
        [MessageBodyMember]
        public int BlockB2bFlag { get; set; }
        [MessageBodyMember]
        public int BlockedFlag { get; set; }
        [MessageBodyMember]
        public int LowComfortFlag { get; set; }
        [MessageBodyMember]
        public int PassengerCount { get; set; }
    }
}
