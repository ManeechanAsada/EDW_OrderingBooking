using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using booking = Avantik.Web.Service.Entity.Booking;

namespace Avantik.Web.Service.Entity.Flight
{
    public class SeatMap : booking.Flight
    {
        public Guid FlightId { get; set; }
        public int FreeSeatingFlag { get; set; }
        public string FlightCheckInStatusRcd { get; set; }
        public string OriginRcd { get; set; }
        public string DestinationRcd { get; set; }
        public string AircraftConfigurationCode { get; set; }
        public int NumberOfBays { get; set; }
        public string BoardingClassRcd { get; set; }
        public int NumberOfRows { get; set; }
        public int NumberOfColumns { get; set; }
        public int LayoutRow { get; set; }
        public int LayoutColumn { get; set; }
        public string LocationTypeRcd { get; set; }
        public string SeatColumn { get; set; }
        public string FeeRcd { get; set; }
        public int SeatRow { get; set; }
        public int StretcherFlag { get; set; }
        public int HanddicappedFlag { get; set; }
        public int NoChildFlag { get; set; }
        public int BassinetFlag { get; set; }
        public int NoInfantFlag { get; set; }
        public int InfantFlag { get; set; }
        public int EmergencyExitFlag { get; set; }
        public int UnAccompaniedMinorsFlag { get; set; }
        public int WindowFlag { get; set; }
        public int AisleFlag { get; set; }
        public int BlockB2cFlag { get; set; }
        public int BlockB2bFlag { get; set; }
        public int BlockedFlag { get; set; }
        public int LowComfortFlag { get; set; }
        public int PassengerCount { get; set; }
    }
}
