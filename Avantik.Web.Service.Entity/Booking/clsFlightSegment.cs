using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Avantik.Web.Service.Entity.Booking
{
    public class FlightSegment : FlightBase
    {
        #region "Property"
        public Guid CreateBy { get; set; }
        public Guid UpdateBy { get; set; }
        public DateTime CreateDateTime { get; set; }
        public DateTime UpdateDateTime { get; set; }
        public Int16 RoutesTot { get; set; }
        public Int16 RoutesAvl { get; set; }
        public Int16 RoutesB2c { get; set; }
        public Int16 RoutesB2b { get; set; }
        public Int16 RoutesB2s { get; set; }
        public Int16 RoutesApi { get; set; }
        public Int16 RoutesB2t { get; set; }

        public bool SegmentChangeFeeFlag { get; set; }
        public bool TransitFlag { get; set; }
        public bool DirectFlag { get; set; }
        public bool AvlFlag { get; set; }
        public bool B2cFlag { get; set; }
        public bool B2bFlag { get; set; }
        public bool B2tFlag { get; set; }
        public Int16 DayRange { get; set; }
        public bool ShowRedressNumberFlag { get; set; }
        public bool RequirePassengerTitleFlag { get; set; }
        public bool RequirePassengerGenderFlag { get; set; }
        public bool RequireDateOfBirthFlag { get; set; }
        public bool RequireDocumentDetailsFlag { get; set; }
        public bool RequirePassengerWeightFlag { get; set; }
        public bool SpecialServiceFeeFlag { get; set; }
        public bool ShowInsuranceOnWebFlag { get; set; }
        public Guid ExchangedSegmentId { get; set; }
        public string FlightCheckInStatusRcd { get; set; }
        public byte RefundableFlag { get; set; }
        public byte GroupFlag { get; set; }
        public byte NonRevenueFlag { get; set; }
        public byte EticketFlag { get; set; }
        public byte FareReduction { get; set; }
        public DateTime ArrivalDate { get; set; }
        public byte WaitlistFlag { get; set; }
        public byte OverbookFlag { get; set; }
        public Guid FlightConnectionId { get; set; }
        public byte AdvancedPurchaseFlag { get; set; }
        public int JourneyTime { get; set; }
        public int DepartureTime { get; set; }
        public int ArrivalTime { get; set; }

        public string OriginName { get; set; }
        public string DestinationName { get; set; }

        public Guid TransitFlightId { get; set; }
        public Guid TransitFareId { get; set; }
        public DateTime TransitDepartureDate { get; set; }
        public DateTime TransitArrivalDate { get; set; }
        public Guid FareId { get; set; }
        public string BookingClassRcd { get; set; }
        public string BoardingClassRcd { get; set; }
        public string AircraftTypeRcd { get; set; }
        public Int16 PlannedDepartureTime { get; set; }
        public Int16 PlannedArrivalTime { get; set; }
        public byte DepartureDayofweek { get; set; }
        public byte ArrivalDayofweek { get; set; }
        public int DepartureMonth { get; set; }
        public Guid BookingSegmentId { get; set; }

        public string SegmentStatusRcd { get; set; }
        public Guid BookingId { get; set; }
        public int NumberOfUnits { get; set; }

        public string SegmentChangeStatusRcd { get; set; }
        public byte InfoSegmentFlag { get; set; }
        public byte HighPriorityWaitlistFlag { get; set; }
        public string PriorityRcd { get; set; }
        public string SegmentStatusName { get; set; }


        public byte SeatmapFlag { get; set; }
        public byte TempSeatmapFlag { get; set; }

        public byte AllowWebCheckinFlag { get; set; }
        public Int16 CloseWebSalesFlag { get; set; }
        public Int16 ExcludeQuoteFlag { get; set; }
        public double CurrencyRate { get; set; }
        public byte OpenSequence { get; set; }
        public byte NumberOfStops { get; set; }

        public Guid FlightId { get; set; }
        public string OdOriginRcd { get; set; }
        public string OdDestinationRcd { get; set; }

        public DateTime FlightFlownDateTime { get; set; }
        public string FlightStatusRCD { get; set; }
        #endregion
    }

    public class SegmentSeatInfo
    {
        public Guid BookingSegmentId { get; set; }
        public string SeatDetails { get; set; }
    }

}
