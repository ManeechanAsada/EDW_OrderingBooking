using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.Serialization;
using System.ServiceModel;

namespace Avantik.Web.Service.Message.Booking
{
    public class FlightSegment
    {
        [MessageBodyMember]
        public string OriginRcd { get; set; }
        [MessageBodyMember]
        public string DestinationRcd { get; set; }
        [MessageBodyMember]
        public Guid CreateBy { get; set; }
        [MessageBodyMember]
        public Guid UpdateBy { get; set; }
        [MessageBodyMember]
        public DateTime CreateDateTime { get; set; }
        [MessageBodyMember]
        public DateTime UpdateDateTime { get; set; }
        [MessageBodyMember]
        public Int16 RoutesTot { get; set; }
        [MessageBodyMember]
        public Int16 RoutesAvl { get; set; }
        [MessageBodyMember]
        public Int16 RoutesB2c { get; set; }
        [MessageBodyMember]
        public Int16 RoutesB2b { get; set; }
        [MessageBodyMember]
        public Int16 RoutesB2s { get; set; }
        [MessageBodyMember]
        public Int16 RoutesApi { get; set; }
        [MessageBodyMember]
        public Int16 RoutesB2t { get; set; }
        [MessageBodyMember]
        public bool SegmentChangeFeeFlag { get; set; }
        [MessageBodyMember]
        public bool TransitFlag { get; set; }
        [MessageBodyMember]
        public bool DirectFlag { get; set; }
        [MessageBodyMember]
        public bool AvlFlag { get; set; }
        [MessageBodyMember]
        public bool B2cFlag { get; set; }
        [MessageBodyMember]
        public bool B2bFlag { get; set; }
        [MessageBodyMember]
        public bool B2tFlag { get; set; }
        [MessageBodyMember]
        public Int16 DayRange { get; set; }
        [MessageBodyMember]
        public bool ShowRedressNumberFlag { get; set; }
        [MessageBodyMember]
        public bool RequirePassengerTitleFlag { get; set; }
        [MessageBodyMember]
        public bool RequirePassengerGenderFlag { get; set; }
        [MessageBodyMember]
        public bool RequireDateOfBirthFlag { get; set; }
        [MessageBodyMember]
        public bool RequireDocumentDetailsFlag { get; set; }
        [MessageBodyMember]
        public bool RequirePassengerWeightFlag { get; set; }
        [MessageBodyMember]
        public bool SpecialServiceFeeFlag { get; set; }
        [MessageBodyMember]
        public bool ShowInsuranceOnWebFlag { get; set; }
        [MessageBodyMember]
        public Guid FlightId { get; set; }
        [MessageBodyMember]
        public Guid ExchangedSegmentId { get; set; }
        [MessageBodyMember]
        public string AirlineRcd { get; set; }
        [MessageBodyMember]
        public string FlightNumber { get; set; }
        [MessageBodyMember]
        public byte RefundableFlag { get; set; }
        [MessageBodyMember]
        public byte GroupFlag { get; set; }
        [MessageBodyMember]
        public byte NonRevenueFlag { get; set; }
        [MessageBodyMember]
        public byte EticketFlag { get; set; }
        [MessageBodyMember]
        public byte FareReduction { get; set; }
        [MessageBodyMember]
        public DateTime DepartureDate { get; set; }
        [MessageBodyMember]
        public DateTime ArrivalDate { get; set; }
        [MessageBodyMember]
        public string OdOriginRcd { get; set; }
        [MessageBodyMember]
        public string OdDestinationRcd { get; set; }
        [MessageBodyMember]
        public byte WaitlistFlag { get; set; }
        [MessageBodyMember]
        public byte OverbookFlag { get; set; }
        [MessageBodyMember]
        public Guid FlightConnectionId { get; set; }
        [MessageBodyMember]
        public byte AdvancedPurchaseFlag { get; set; }
        [MessageBodyMember]
        public int JourneyTime { get; set; }
        [MessageBodyMember]
        public int DepartureTime { get; set; }
        [MessageBodyMember]
        public int ArrivalTime { get; set; }
        [MessageBodyMember]
        public string OriginName { get; set; }
        [MessageBodyMember]
        public string DestinationName { get; set; }
        [MessageBodyMember]
        public Guid TransitFlightId { get; set; }
        [MessageBodyMember]
        public Guid TransitFareId { get; set; }
        [MessageBodyMember]
        public DateTime TransitDepartureDate { get; set; }
        [MessageBodyMember]
        public DateTime TransitArrivalDate { get; set; }
        [MessageBodyMember]
        public Guid FareId { get; set; }
        [MessageBodyMember]
        public string BookingClassRcd { get; set; }
        [MessageBodyMember]
        public string BoardingClassRcd { get; set; }
        [MessageBodyMember]
        public string AircraftTypeRcd { get; set; }
        [MessageBodyMember]
        public Int16 PlannedDepartureTime { get; set; }
        [MessageBodyMember]
        public Int16 PlannedArrivalTime { get; set; }
        [MessageBodyMember]
        public byte DepartureDayofweek { get; set; }
        [MessageBodyMember]
        public byte ArrivalDayofweek { get; set; }
        [MessageBodyMember]
        public int DepartureMonth { get; set; }
        [MessageBodyMember]
        public Guid BookingSegmentId { get; set; }
        [MessageBodyMember]
        public string SegmentStatusRcd { get; set; }
        [MessageBodyMember]
        public Guid BookingId { get; set; }
        [MessageBodyMember]
        public int NumberOfUnits { get; set; }
        [MessageBodyMember]
        public string SegmentChangeStatusRcd { get; set; }
        [MessageBodyMember]
        public byte InfoSegmentFlag { get; set; }
        [MessageBodyMember]
        public byte HighPriorityWaitlistFlag { get; set; }
        [MessageBodyMember]
        public string SegmentStatusName { get; set; }
        [MessageBodyMember]
        public byte SeatmapFlag { get; set; }
        [MessageBodyMember]
        public byte TempSeatmapFlag { get; set; }
        [MessageBodyMember]
        public byte AllowWebCheckinFlag { get; set; }
        [MessageBodyMember]
        public Int16 CloseWebSalesFlag { get; set; }
        [MessageBodyMember]
        public Int16 ExcludeQuoteFlag { get; set; }
        [MessageBodyMember]
        public double CurrencyRate { get; set; }
        [MessageBodyMember]
        public byte OpenSequence { get; set; }
        [MessageBodyMember]
        public byte NumberOfStops { get; set; }
        [MessageBodyMember]
        public string FlightCheckInStatusRcd { get; set; }
        [MessageBodyMember]
        public DateTime FlightFlownDateTime { get; set; }
        [MessageBodyMember]
        public string FlightStatusRCD { get; set; }

        
    }
}
