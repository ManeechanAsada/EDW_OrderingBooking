using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Avantik.Web.Service.Entity.Booking
{
    //checkin API
    public class Mapping
    {
        public string OriginRcd { get; set; }
        public string DestinationRcd { get; set; }
        public string DisplayName { get; set; }
        public Guid CreateBy { get; set; }
        public Guid UpdateBy { get; set; }
        public DateTime CreateDateTime { get; set; }
        public DateTime UpdateDateTime { get; set; }

        #region Origin Flag
        public string CurrencyRcd { get; set; }
        public Int16 RoutesTot { get; set; }
        public Int16 RoutesAvl { get; set; }
        public Int16 RoutesB2c { get; set; }
        public Int16 RoutesB2b { get; set; }
        public Int16 RoutesB2s { get; set; }
        public Int16 RoutesApi { get; set; }
        public Int16 RoutesB2t { get; set; }
        #endregion

        #region Destination Flag
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
        #endregion

        #region "Property"
        public Guid FlightId { get; set; }
        public Guid ExchangedSegmentId { get; set; }
        public string AirlineRcd { get; set; }
        public string FlightNumber { get; set; }
        public byte RefundableFlag { get; set; }
        public byte GroupFlag { get; set; }
        public byte NonRevenueFlag { get; set; }
        public byte EticketFlag { get; set; }
        public byte FareReduction { get; set; }
        public DateTime DepartureDate { get; set; }
        public DateTime ArrivalDate { get; set; }
        public string OdOriginRcd { get; set; }
        public string OdDestinationRcd { get; set; }
        public byte WaitlistFlag { get; set; }
        public byte OverbookFlag { get; set; }
        public Guid FlightConnectionId { get; set; }
        public byte AdvancedPurchaseFlag { get; set; }
        public int JourneyTime { get; set; }
        public int DepartureTime { get; set; }
        public int ArrivalTime { get; set; }

        public Guid TransitFlightId { get; set; }
        public Guid TransitFareId { get; set; }
        public DateTime TransitDepartureDate { get; set; }
        public DateTime TransitArrivalDate { get; set; }
        public Guid FareId { get; set; }
        public string BookingClassRcd { get; set; }
        public string BoardingClassRcd { get; set; }

        public Int16 PlannedDepartureTime { get; set; }
        public Int16 PlannedArrivalTime { get; set; }
        public byte DepartureDayofweek { get; set; }
        public byte ArrivalDayofweek { get; set; }
        public int DepartureMonth { get; set; }
        #endregion

        #region "Property"

        public Guid BookingSegmentId { get; set; }
        public Guid BookingId { get; set; }
        public int NumberOfUnits { get; set; }
        public byte InfoSegmentFlag { get; set; }
        public byte HighPriorityWaitlistFlag { get; set; }
        public byte SeatmapFlag { get; set; }
        public byte TempSeatmapFlag { get; set; }
        public byte AllowWebCheckinFlag { get; set; }
        public Int16 CloseWebSalesFlag { get; set; }
        public Int16 ExcludeQuoteFlag { get; set; }
        public double CurrencyRate { get; set; }
        public byte OpenSequence { get; set; }
        public byte NumberOfStops { get; set; }
        public string PassengerCheckinStatusRcd { get; set; }
        #endregion

        #region Passenger information
        public Guid PassengerId { get; set; }
        public string Lastname { get; set; }
        public string Firstname { get; set; }
        public string GenderTypeRcd { get; set; }
        public string TitleRcd { get; set; }
        public string PassengerTypeRcd { get; set; }
        public DateTime DateOfBirth { get; set; }
        public string PassengerStatusRcd { get; set; }
        public string Middlename { get; set; }
        public string DocumentTypeRcd { get; set; }
        public string PassportNumber { get; set; }
        public string PassportIssuePlace { get; set; }
        public DateTime PassportIssueDate { get; set; }
        public DateTime PassportExpiryDate { get; set; }
        public string Sendmail { get; set; }
        public long ClientNumber { get; set; }
        public Guid PassengerProfileId { get; set; }
        public Guid ClientProfileId { get; set; }
        public string EmployeeNumber { get; set; }
        public string NationalityRcd { get; set; }
        public string ContactEmail { get; set; }

        #endregion

        #region Ticket information
        public string InventoryClassRcd { get; set; }
        public string SeatNumber { get; set; }
        public int SeatRow { get; set; }
        public string SeatColumn { get; set; }
        public decimal NetTotal { get; set; }
        public decimal TaxAmount { get; set; }
        public decimal FareAmount { get; set; }
        public decimal YqAmount { get; set; }
        public decimal BaseTicketingFeeAmount { get; set; }
        public decimal TicketingFeeAmount { get; set; }
        public decimal ReservationFeeAmount { get; set; }
        public decimal CommissionAmount { get; set; }
        public decimal FareVat { get; set; }
        public decimal TaxVat { get; set; }
        public decimal YqVat { get; set; }
        public decimal TicketingFeeVat { get; set; }
        public decimal ReservationFeeVat { get; set; }
        public decimal FareAmountIncl { get; set; }
        public decimal TaxAmountIncl { get; set; }
        public decimal YqAmountIncl { get; set; }
        public decimal PublicFareAmountIncl { get; set; }
        public decimal PublicFareAmount { get; set; }
        public decimal CommissionAmountIncl { get; set; }
        public decimal CommissionPercentage { get; set; }
        public decimal TicketingFeeAmountIncl { get; set; }
        public decimal ReservationFeeAmountIncl { get; set; }
        public decimal AcctNetTotal { get; set; }
        public decimal AcctTaxAmount { get; set; }
        public decimal AcctFareAmount { get; set; }
        public decimal AcctYqAmount { get; set; }
        public decimal AcctTicketingFeeAmount { get; set; }
        public decimal AcctReservationFeeAmount { get; set; }
        public decimal AcctCommissionAmount { get; set; }
        public decimal AcctFareAmountIncl { get; set; }
        public decimal AcctTaxAmountIncl { get; set; }
        public decimal AcctYqAmountIncl { get; set; }
        public decimal AcctCommissionAmountIncl { get; set; }
        public decimal AcctTicketingFeeAmountIncl { get; set; }
        public decimal AcctReservationFeeAmountIncl { get; set; }
        public string FareCode { get; set; }
        public DateTime FareDateTime { get; set; }
        public byte ETicketFlag { get; set; }
        public byte StandbyFlag { get; set; }
        public byte TransferableFareFlag { get; set; }
        public string AgencyCode { get; set; }
        public string TicketNumber { get; set; }
        public DateTime TicketingDateTime { get; set; }
        public Guid TicketingBy { get; set; }
        public int CheckInSequence { get; set; }
        public int GroupSequence { get; set; }
        public Guid UnloadBy { get; set; }
        public DateTime UnloadDateTime { get; set; }
        public int NumberOfPieces { get; set; }
        public decimal BaggageWeight { get; set; }
        public decimal ExcessBaggageWeight { get; set; }
        public decimal CheckInBaggageWeight { get; set; }
        public Guid CheckInBy { get; set; }
        public Guid CancelBy { get; set; }
        public DateTime ExchangedDateTime { get; set; }
        public DateTime CancelDateTime { get; set; }
        public string RestrictionText { get; set; }
        public string EndorsementText { get; set; }
        public byte ExcludePricingFlag { get; set; }
        public string CreateName { get; set; }
        public string UpdateName { get; set; }
        public DateTime VoidDateTime { get; set; }
        public decimal RefundCharge { get; set; }
        public decimal RefundableAmount { get; set; }
        public DateTime RefundDateTime { get; set; }
        public decimal PaymentAmount { get; set; }
        public int MappingSort { get; set; }
        public DateTime CheckInDateTime { get; set; }
        public DateTime OnwardDepartureDate { get; set; }
        public string ETicketStatus { get; set; }
        public byte HandLuggageFlag { get; set; }
        public int HandNumberOfPieces { get; set; }
        public double HandBaggageWeight { get; set; }
        public byte FreeSeatingFlag { get; set; }
        public string FareTypeRcd { get; set; }
        public double RedemptionPoints { get; set; }
        public string SeatFeeRcd { get; set; }
        public Int16 RefundWithChargeHours { get; set; }
        public Int16 RefundNotPossibleHours { get; set; }
        public byte DutyTravelFlag { get; set; }
        public DateTime NotValidAfterDate { get; set; }
        public DateTime NotValidBeforeDate { get; set; }
        public byte AdvancedSeatingFlag { get; set; }
        public byte FareColumn { get; set; }
        public decimal ExchangedPaid { get; set; }
        public decimal TransferableAmount { get; set; }

        #endregion

        public Int16 PieceAllowance { get; set; }
        public int BoardingTime { get; set; }
        public byte ItFareFlag { get; set; }

        public byte ThroughFareFlag { get; set; }

        #region Method
        public PaymentAllocation GetAllocation(Guid bookingPaymentId, 
                                               Guid userId)
        {
            decimal allocationAmount = 0;
            if (ExcludePricingFlag == 1)
            {
                allocationAmount = 0;
            }
            else if (RefundDateTime != DateTime.MinValue)
            {
                allocationAmount = RefundCharge - PaymentAmount;
            }
            else
            {
                allocationAmount = NetTotal - PaymentAmount - ExchangedPaid;
            }

            if (allocationAmount > 0)
            {
                PaymentAllocation allocation = new PaymentAllocation();

                allocation.booking_payment_id = bookingPaymentId;
                allocation.passenger_id = PassengerId;
                allocation.booking_segment_id = BookingSegmentId;
                allocation.user_id = userId;

                allocation.currency_rcd = CurrencyRcd;
                allocation.charge_currency_rcd = CurrencyRcd;
                allocation.od_origin_rcd = OdOriginRcd;
                allocation.od_destination_rcd = OdDestinationRcd;

                allocation.sales_amount = allocationAmount;

                allocation.payment_amount = PaymentAmount;

                return allocation;
            }
            else
            {
                return null;
            }
            
        }
        public bool GetExchangeTicket(Guid bookingSegmentId, Guid passengerId)
        {
            if (ExchangedDateTime != DateTime.MinValue &&
                ExchangedSegmentId == bookingSegmentId &&
                PassengerId == passengerId)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        #endregion
    }
}
