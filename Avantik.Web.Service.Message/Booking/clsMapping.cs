using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.Serialization;
using System.ServiceModel;

namespace Avantik.Web.Service.Message.Booking
{
    public class Mapping 
    {
        public string OriginRcd { get; set; }
 	 	 [MessageBodyMember] 
        public string DestinationRcd { get; set; }
 	 	 [MessageBodyMember] 
        public string DisplayName { get; set; }
 	 	 [MessageBodyMember] 
        public Guid CreateBy { get; set; }
 	 	 [MessageBodyMember] 
        public Guid UpdateBy { get; set; }
         [MessageBodyMember]
        public DateTime CreateDateTime { get; set; }
         [MessageBodyMember]
        public DateTime UpdateDateTime { get; set; }
 	 	 [MessageBodyMember] 
        public string CurrencyRcd { get; set; }
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
 	 	 [DataMember(EmitDefaultValue = false)]
        public DateTime DepartureDate { get; set; }
 	 	 [DataMember(EmitDefaultValue = false)]
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
        public Guid BookingId { get; set; }
 	 	 [MessageBodyMember] 
        public int NumberOfUnits { get; set; }
 	 	 [MessageBodyMember] 
        public byte InfoSegmentFlag { get; set; }
 	 	 [MessageBodyMember] 
        public byte HighPriorityWaitlistFlag { get; set; }
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
        public Guid PassengerId { get; set; }
 	 	 [MessageBodyMember] 
        public string Lastname { get; set; }
 	 	 [MessageBodyMember] 
        public string Firstname { get; set; }
 	 	 [MessageBodyMember] 
        public string GenderTypeRcd { get; set; }
 	 	 [MessageBodyMember] 
        public string TitleRcd { get; set; }
 	 	 [MessageBodyMember] 
        public string PassengerTypeRcd { get; set; }
         [MessageBodyMember]
        public DateTime DateOfBirth { get; set; }
 	 	 [MessageBodyMember] 
        public string PassengerStatusRcd { get; set; }
 	 	 [MessageBodyMember] 
        public string Middlename { get; set; }
 	 	 [MessageBodyMember] 
        public string DocumentTypeRcd { get; set; }
 	 	 [MessageBodyMember] 
        public string PassportNumber { get; set; }
 	 	 [MessageBodyMember] 
        public string PassportIssuePlace { get; set; }
         [MessageBodyMember]
        public DateTime PassportIssueDate { get; set; }
         [MessageBodyMember]
        public DateTime PassportExpiryDate { get; set; }
 	 	 [MessageBodyMember] 
        public string Sendmail { get; set; }
 	 	 [MessageBodyMember] 
        public long ClientNumber { get; set; }
 	 	 [MessageBodyMember] 
        public Guid PassengerProfileId { get; set; }
 	 	 [MessageBodyMember] 
        public Guid ClientProfileId { get; set; }
 	 	 [MessageBodyMember] 
        public string EmployeeNumber { get; set; }
 	 	 [MessageBodyMember] 
        public string NationalityRcd { get; set; }
 	 	 [MessageBodyMember] 
        public string ContactEmail { get; set; }
 	 	 [MessageBodyMember] 
        public string InventoryClassRcd { get; set; }
 	 	 [MessageBodyMember] 
        public string SeatNumber { get; set; }
 	 	 [MessageBodyMember] 
        public int SeatRow { get; set; }
 	 	 [MessageBodyMember] 
        public string SeatColumn { get; set; }
 	 	 [MessageBodyMember] 
        public decimal NetTotal { get; set; }
 	 	 [MessageBodyMember] 
        public decimal TaxAmount { get; set; }
 	 	 [MessageBodyMember] 
        public decimal FareAmount { get; set; }
 	 	 [MessageBodyMember] 
        public decimal YqAmount { get; set; }
 	 	 [MessageBodyMember] 
        public decimal BaseTicketingFeeAmount { get; set; }
 	 	 [MessageBodyMember] 
        public decimal TicketingFeeAmount { get; set; }
 	 	 [MessageBodyMember] 
        public decimal ReservationFeeAmount { get; set; }
 	 	 [MessageBodyMember] 
        public decimal CommissionAmount { get; set; }
 	 	 [MessageBodyMember] 
        public decimal FareVat { get; set; }
 	 	 [MessageBodyMember] 
        public decimal TaxVat { get; set; }
 	 	 [MessageBodyMember] 
        public decimal YqVat { get; set; }
 	 	 [MessageBodyMember] 
        public decimal TicketingFeeVat { get; set; }
 	 	 [MessageBodyMember] 
        public decimal ReservationFeeVat { get; set; }
 	 	 [MessageBodyMember] 
        public decimal FareAmountIncl { get; set; }
 	 	 [MessageBodyMember] 
        public decimal TaxAmountIncl { get; set; }
 	 	 [MessageBodyMember] 
        public decimal YqAmountIncl { get; set; }
 	 	 [MessageBodyMember] 
        public decimal PublicFareAmountIncl { get; set; }
 	 	 [MessageBodyMember] 
        public decimal PublicFareAmount { get; set; }
 	 	 [MessageBodyMember] 
        public decimal CommissionAmountIncl { get; set; }
 	 	 [MessageBodyMember] 
        public decimal CommissionPercentage { get; set; }
 	 	 [MessageBodyMember] 
        public decimal TicketingFeeAmountIncl { get; set; }
 	 	 [MessageBodyMember] 
        public decimal ReservationFeeAmountIncl { get; set; }
 	 	 [MessageBodyMember] 
        public decimal AcctNetTotal { get; set; }
 	 	 [MessageBodyMember] 
        public decimal AcctTaxAmount { get; set; }
 	 	 [MessageBodyMember] 
        public decimal AcctFareAmount { get; set; }
 	 	 [MessageBodyMember] 
        public decimal AcctYqAmount { get; set; }
 	 	 [MessageBodyMember] 
        public decimal AcctTicketingFeeAmount { get; set; }
 	 	 [MessageBodyMember] 
        public decimal AcctReservationFeeAmount { get; set; }
 	 	 [MessageBodyMember] 
        public decimal AcctCommissionAmount { get; set; }
 	 	 [MessageBodyMember] 
        public decimal AcctFareAmountIncl { get; set; }
 	 	 [MessageBodyMember] 
        public decimal AcctTaxAmountIncl { get; set; }
 	 	 [MessageBodyMember] 
        public decimal AcctYqAmountIncl { get; set; }
 	 	 [MessageBodyMember] 
        public decimal AcctCommissionAmountIncl { get; set; }
 	 	 [MessageBodyMember] 
        public decimal AcctTicketingFeeAmountIncl { get; set; }
 	 	 [MessageBodyMember] 
        public decimal AcctReservationFeeAmountIncl { get; set; }
 	 	 [MessageBodyMember] 
        public string FareCode { get; set; }
         [MessageBodyMember]
        public DateTime FareDateTime { get; set; }
 	 	 [MessageBodyMember] 
        public byte ETicketFlag { get; set; }
 	 	 [MessageBodyMember] 
        public byte StandbyFlag { get; set; }
 	 	 [MessageBodyMember] 
        public byte TransferableFareFlag { get; set; }
 	 	 [MessageBodyMember] 
        public string AgencyCode { get; set; }
 	 	 [MessageBodyMember] 
        public string TicketNumber { get; set; }
         [MessageBodyMember]
        public DateTime TicketingDateTime { get; set; }
 	 	 [MessageBodyMember] 
        public Guid TicketingBy { get; set; }
 	 	 [MessageBodyMember] 
        public int CheckInSequence { get; set; }
 	 	 [MessageBodyMember] 
        public int GroupSequence { get; set; }
 	 	 [MessageBodyMember] 
        public Guid UnloadBy { get; set; }
         [MessageBodyMember]
        public DateTime UnloadDateTime { get; set; }
 	 	 [MessageBodyMember] 
        public int NumberOfPieces { get; set; }
 	 	 [MessageBodyMember] 
        public decimal BaggageWeight { get; set; }
 	 	 [MessageBodyMember] 
        public decimal ExcessBaggageWeight { get; set; }
 	 	 [MessageBodyMember] 
        public decimal CheckInBaggageWeight { get; set; }
 	 	 [MessageBodyMember] 
        public Guid CheckInBy { get; set; }
 	 	 [MessageBodyMember] 
        public Guid CancelBy { get; set; }
         [MessageBodyMember]
        public DateTime ExchangedDateTime { get; set; }
         [MessageBodyMember]
        public DateTime CancelDateTime { get; set; }
 	 	 [MessageBodyMember] 
        public string RestrictionText { get; set; }
 	 	 [MessageBodyMember] 
        public string EndorsementText { get; set; }
 	 	 [MessageBodyMember] 
        public byte ExcludePricingFlag { get; set; }
 	 	 [MessageBodyMember] 
        public string CreateName { get; set; }
 	 	 [MessageBodyMember] 
        public string UpdateName { get; set; }
         [MessageBodyMember]
        public DateTime VoidDateTime { get; set; }
 	 	 [MessageBodyMember] 
        public decimal RefundCharge { get; set; }
 	 	 [MessageBodyMember] 
        public decimal RefundableAmount { get; set; }
 	 	 [DataMember(EmitDefaultValue = false)]
        public DateTime RefundDateTime { get; set; }
 	 	 [MessageBodyMember] 
        public decimal PaymentAmount { get; set; }
 	 	 [MessageBodyMember] 
        public int MappingSort { get; set; }
         [MessageBodyMember]
        public DateTime CheckInDateTime { get; set; }
         [MessageBodyMember]
        public DateTime OnwardDepartureDate { get; set; }
 	 	 [MessageBodyMember] 
        public string ETicketStatus { get; set; }
 	 	 [MessageBodyMember] 
        public byte HandLuggageFlag { get; set; }
 	 	 [MessageBodyMember] 
        public int HandNumberOfPieces { get; set; }
 	 	 [MessageBodyMember] 
        public double HandBaggageWeight { get; set; }
 	 	 [MessageBodyMember] 
        public byte FreeSeatingFlag { get; set; }
 	 	 [MessageBodyMember] 
        public string FareTypeRcd { get; set; }
 	 	 [MessageBodyMember] 
        public double RedemptionPoints { get; set; }
 	 	 [MessageBodyMember] 
        public string SeatFeeRcd { get; set; }
 	 	 [MessageBodyMember] 
        public Int16 RefundWithChargeHours { get; set; }
 	 	 [MessageBodyMember] 
        public Int16 RefundNotPossibleHours { get; set; }
 	 	 [MessageBodyMember] 
        public byte DutyTravelFlag { get; set; }
         [MessageBodyMember]
        public DateTime NotValidAfterDate { get; set; }
         [MessageBodyMember]
        public DateTime NotValidBeforeDate { get; set; }
 	 	 [MessageBodyMember] 
        public byte AdvancedSeatingFlag { get; set; }
 	 	 [MessageBodyMember] 
        public byte FareColumn { get; set; }
 	 	 [MessageBodyMember] 
        public Int16 PieceAllowance { get; set; }
 	 	 [MessageBodyMember] 
        public int BoardingTime { get; set; }
 	 	 [MessageBodyMember] 
        public byte ItFareFlag { get; set; }
        [MessageBodyMember]
        public byte ThroughFareFlag { get; set; }
        [MessageBodyMember]
        public string PassengerCheckinStatusRcd { get; set; }
        

    }
}
