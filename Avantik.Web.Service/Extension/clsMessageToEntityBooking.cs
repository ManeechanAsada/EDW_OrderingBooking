using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Avantik.Web.Service.Entity.Booking;

namespace Avantik.Web.Service.Extension
{
    public static class MessageToEntityBooking
    {
        public static BookingHeader ToEntity(this  Message.Booking.BookingHeader header)
        {
            BookingHeader bookingHeader = null;
            if (header != null)
            {
                bookingHeader = new BookingHeader();

                bookingHeader.BookingId = header.BookingId;

                if (header.OwnAgencyFlag == 1 && header.WebAgencyFlag == 1)
                    bookingHeader.BookingSourceRcd = "B2C";
                else if (header.OwnAgencyFlag == 0 && header.WebAgencyFlag == 1)
                    bookingHeader.BookingSourceRcd = "B2B";

                bookingHeader.CurrencyRcd = header.CurrencyRcd;
                bookingHeader.ClientProfileId = header.ClientProfileId;
                bookingHeader.BookingNumber = header.BookingNumber;
                bookingHeader.RecordLocator = header.RecordLocator;
                bookingHeader.NumberOfAdults = header.NumberOfAdults;
                bookingHeader.NumberOfChildren = header.NumberOfChildren;
                bookingHeader.NumberOfInfants = header.NumberOfInfants;

                if (!String.IsNullOrEmpty(header.LanguageRcd))
                {
                    bookingHeader.LanguageRcd = header.LanguageRcd;
                }
                else
                {
                    bookingHeader.LanguageRcd = "EN";
                }

                bookingHeader.AgencyCode = header.AgencyCode;
                bookingHeader.ContactName = header.ContactName;
                bookingHeader.ContactEmail = header.ContactEmail;
                bookingHeader.PhoneMobile = header.PhoneMobile;
                bookingHeader.PhoneHome = header.PhoneHome;
                bookingHeader.PhoneBusiness = header.PhoneBusiness;
                bookingHeader.ReceivedFrom = header.ReceivedFrom;
                bookingHeader.PhoneFax = header.PhoneFax;
                bookingHeader.PhoneSearch = header.PhoneSearch;
                bookingHeader.Comment = header.Comment;
                bookingHeader.NotifyByEmailFlag = header.NotifyByEmailFlag;
                bookingHeader.NotifyBySmsFlag = header.NotifyBySmsFlag;
                bookingHeader.GroupName = header.GroupName;
                bookingHeader.GroupBookingFlag = header.GroupBookingFlag;
                bookingHeader.AgencyName = header.AgencyName;
                bookingHeader.OwnAgencyFlag = header.OwnAgencyFlag;
                bookingHeader.WebAgencyFlag = header.WebAgencyFlag;
                bookingHeader.ClientNumber = header.ClientNumber;
                bookingHeader.Lastname = header.Lastname;
                bookingHeader.Firstname = header.Firstname;
                bookingHeader.City = header.City;
                bookingHeader.CreateName = header.CreateName;
                bookingHeader.Street = header.Street;
                bookingHeader.LockDateTime = header.LockDateTime;
                bookingHeader.Middlename = header.Middlename;
                bookingHeader.AddressLine1 = header.AddressLine1;
                bookingHeader.AddressLine2 = header.AddressLine2;
                bookingHeader.State = header.State;
                bookingHeader.District = header.District;
                bookingHeader.Province = header.Province;
                bookingHeader.ZipCode = header.ZipCode;
                bookingHeader.PoBox = header.PoBox;
                bookingHeader.CountryRcd = header.CountryRcd;
                bookingHeader.TitleRcd = header.TitleRcd;
                bookingHeader.ExternalPaymentReference = header.ExternalPaymentReference;
                bookingHeader.CreateBy = header.CreateBy;
                bookingHeader.UpdateBy = header.UpdateBy;
                bookingHeader.CreateDateTime = header.CreateDateTime;
                bookingHeader.UpdateDateTime = header.UpdateDateTime;
                bookingHeader.CostCenter = header.CostCenter;
                bookingHeader.PurchaseOrder = header.PurchaseOrder;
                bookingHeader.ProjectNumber = header.ProjectNumber;
                bookingHeader.LockName = header.LockName;
                bookingHeader.IpAddress = header.IpAddress;
                bookingHeader.ApprovalFlag = header.ApprovalFlag;
                bookingHeader.InvoiceReceiver = header.InvoiceReceiver;
                bookingHeader.TaxId = header.TaxId;
                bookingHeader.NewsletterFlag = header.NewsletterFlag;
                bookingHeader.ContactEmailCc = header.ContactEmailCc;
                bookingHeader.MobileEmail = header.MobileEmail;
                bookingHeader.VendorRcd = header.VendorRcd;
                bookingHeader.BookingDateTime = header.BookingDateTime;
                bookingHeader.NoVatFlag = header.NoVatFlag;
                bookingHeader.CompanyName = header.CompanyName;
                bookingHeader.BusinessFlag = header.BusinessFlag;

            }

            return bookingHeader;
        }
        public static IList<FlightSegment> ToListEntity(this  IList<Message.Booking.FlightSegment> segment)
        {
            IList<FlightSegment> segmentList = null;
            if (segment != null)
            {
                segmentList = new List<FlightSegment>();

                for (int i = 0; i < segment.Count; i++)
                {
                    segmentList.Add(segment[i].ToEntity());
                }
            }

            return segmentList;
        }
        public static FlightSegment ToEntity(this  Message.Booking.FlightSegment segment)
        {
            FlightSegment flightSegment = null;
            if (segment != null)
            {
                flightSegment = new FlightSegment();
                flightSegment.OriginRcd = segment.OriginRcd;
                flightSegment.DestinationRcd = segment.DestinationRcd;
                flightSegment.CreateBy = segment.CreateBy;
                flightSegment.UpdateBy = segment.UpdateBy;
                flightSegment.CreateDateTime = segment.CreateDateTime;
                flightSegment.UpdateDateTime = segment.UpdateDateTime;
                flightSegment.RoutesTot = segment.RoutesTot;
                flightSegment.RoutesAvl = segment.RoutesAvl;
                flightSegment.RoutesB2c = segment.RoutesB2c;
                flightSegment.RoutesB2b = segment.RoutesB2b;
                flightSegment.RoutesB2s = segment.RoutesB2s;
                flightSegment.RoutesApi = segment.RoutesApi;
                flightSegment.RoutesB2t = segment.RoutesB2t;

                flightSegment.SegmentChangeFeeFlag = segment.SegmentChangeFeeFlag;
                flightSegment.TransitFlag = segment.TransitFlag;
                flightSegment.DirectFlag = segment.DirectFlag;
                flightSegment.AvlFlag = segment.AvlFlag;
                flightSegment.B2cFlag = segment.B2cFlag;
                flightSegment.B2bFlag = segment.B2bFlag;
                flightSegment.B2tFlag = segment.B2tFlag;
                flightSegment.DayRange = segment.DayRange;
                flightSegment.ShowRedressNumberFlag = segment.ShowRedressNumberFlag;
                flightSegment.RequirePassengerTitleFlag = segment.RequirePassengerTitleFlag;
                flightSegment.RequirePassengerGenderFlag = segment.RequirePassengerGenderFlag;
                flightSegment.RequireDateOfBirthFlag = segment.RequireDateOfBirthFlag;
                flightSegment.RequireDocumentDetailsFlag = segment.RequireDocumentDetailsFlag;
                flightSegment.RequirePassengerWeightFlag = segment.RequirePassengerWeightFlag;
                flightSegment.SpecialServiceFeeFlag = segment.SpecialServiceFeeFlag;
                flightSegment.ShowInsuranceOnWebFlag = segment.ShowInsuranceOnWebFlag;
                flightSegment.FlightId = segment.FlightId;
                flightSegment.ExchangedSegmentId = segment.ExchangedSegmentId;
                flightSegment.AirlineRcd = segment.AirlineRcd;
                flightSegment.FlightNumber = segment.FlightNumber;
                flightSegment.RefundableFlag = segment.RefundableFlag;
                flightSegment.GroupFlag = segment.GroupFlag;
                flightSegment.NonRevenueFlag = segment.NonRevenueFlag;
                flightSegment.EticketFlag = segment.EticketFlag;
                flightSegment.FareReduction = segment.FareReduction;
                flightSegment.DepartureDate = segment.DepartureDate;
                flightSegment.ArrivalDate = segment.ArrivalDate;
                flightSegment.OdOriginRcd = segment.OdOriginRcd;
                flightSegment.OdDestinationRcd = segment.OdDestinationRcd;
                flightSegment.WaitlistFlag = segment.WaitlistFlag;
                flightSegment.OverbookFlag = segment.OverbookFlag;
                flightSegment.FlightConnectionId = segment.FlightConnectionId;
                flightSegment.AdvancedPurchaseFlag = segment.AdvancedPurchaseFlag;
                flightSegment.JourneyTime = segment.JourneyTime;
                flightSegment.DepartureTime = segment.DepartureTime;
                flightSegment.ArrivalTime = segment.ArrivalTime;
                flightSegment.ArrivalDate = segment.ArrivalDate;
                flightSegment.OriginName = segment.OriginName;
                flightSegment.DestinationName = segment.DestinationName;

                flightSegment.TransitFlightId = segment.TransitFlightId;
                flightSegment.TransitFareId = segment.TransitFareId;
                flightSegment.TransitDepartureDate = segment.TransitDepartureDate;
                flightSegment.TransitArrivalDate = segment.TransitArrivalDate;
                flightSegment.FareId = segment.FareId;
                flightSegment.BookingClassRcd = segment.BookingClassRcd;
                flightSegment.BoardingClassRcd = segment.BoardingClassRcd;
                flightSegment.AircraftTypeRcd = segment.AircraftTypeRcd;
                flightSegment.PlannedDepartureTime = segment.PlannedDepartureTime;
                flightSegment.PlannedArrivalTime = segment.PlannedArrivalTime;
                flightSegment.DepartureDayofweek = segment.DepartureDayofweek;
                flightSegment.ArrivalDayofweek = segment.ArrivalDayofweek;
                flightSegment.DepartureMonth = segment.DepartureMonth;
                flightSegment.BookingSegmentId = segment.BookingSegmentId;

                flightSegment.SegmentStatusRcd = segment.SegmentStatusRcd;

                flightSegment.BookingId = segment.BookingId;

                flightSegment.NumberOfUnits = segment.NumberOfUnits;

                flightSegment.SegmentChangeStatusRcd = segment.SegmentChangeStatusRcd;
                flightSegment.InfoSegmentFlag = segment.InfoSegmentFlag;
                flightSegment.HighPriorityWaitlistFlag = segment.HighPriorityWaitlistFlag;
                flightSegment.SegmentStatusName = segment.SegmentStatusName;

                flightSegment.SeatmapFlag = segment.SeatmapFlag;
                flightSegment.TempSeatmapFlag = segment.TempSeatmapFlag;

                flightSegment.AllowWebCheckinFlag = segment.AllowWebCheckinFlag;
                flightSegment.CloseWebSalesFlag = segment.CloseWebSalesFlag;
                flightSegment.ExcludeQuoteFlag = segment.ExcludeQuoteFlag;
                flightSegment.CurrencyRate = segment.CurrencyRate;
                flightSegment.OpenSequence = segment.OpenSequence;
                flightSegment.NumberOfStops = segment.NumberOfStops;
                flightSegment.FlightCheckInStatusRcd = segment.FlightCheckInStatusRcd;
                flightSegment.FlightFlownDateTime = segment.FlightFlownDateTime;
                flightSegment.FlightStatusRCD = segment.FlightStatusRCD;

                flightSegment.ExchangedSegmentId = segment.ExchangedSegmentId;

            }

            return flightSegment;
        }
        public static IList<Passenger> ToListEntity(this  IList<Message.Booking.Passenger> p)
        {
            IList<Passenger> passengerList = null;

            if (p != null)
            {
                passengerList = new List<Passenger>();
                for (int i = 0; i < p.Count; i++)
                {
                    passengerList.Add(p[i].ToEntity());
                }
            }

            return passengerList;
        }
        public static Passenger ToEntity(this  Message.Booking.Passenger p)
        {
            Passenger passenger = null;
            if (p != null)
            {
                passenger = new Passenger();

                passenger.TitleRcd = p.TitleRcd;
                passenger.Lastname = p.Lastname;
                passenger.Firstname = p.Firstname;
                passenger.Middlename = p.Middlename;
                passenger.NationalityRcd = p.NationalityRcd;
                passenger.PassengerWeight = p.PassengerWeight;
                passenger.GenderTypeRcd = p.GenderTypeRcd;
                passenger.PassengerTypeRcd = p.PassengerTypeRcd;
                passenger.ClientProfileId = p.ClientProfileId;
                passenger.ClientNumber = p.ClientNumber;

                passenger.AddressLine1 = p.AddressLine1;
                passenger.AddressLine2 = p.AddressLine2;
                passenger.State = p.State;
                passenger.District = p.District;
                passenger.Province = p.Province;
                passenger.ZipCode = p.ZipCode;
                passenger.PoBox = p.PoBox;
                passenger.CountryRcd = p.CountryRcd;
                passenger.Street = p.Street;
                passenger.City = p.City;

                passenger.DocumentTypeRcd = p.DocumentTypeRcd;
                passenger.DocumentNumber = p.DocumentNumber;
                passenger.ResidenceCountryRcd = p.ResidenceCountryRcd;
                passenger.PassportNumber = p.PassportNumber;
                passenger.PassportIssueDate = p.PassportIssueDate;
                passenger.PassportExpiryDate = p.PassportExpiryDate;
                passenger.PassportIssuePlace = p.PassportIssuePlace;
                passenger.PassportBirthPlace = p.PassportBirthPlace;
                passenger.DateOfBirth = p.DateOfBirth;
                passenger.PassportIssueCountryRcd = p.PassportIssueCountryRcd;
                passenger.CountryCodeLong = p.CountryCodeLong;

                passenger.ContactName = p.ContactName;
                passenger.ContactEmail = p.ContactEmail;
                passenger.MobileEmail = p.MobileEmail;
                passenger.PhoneMobile = p.PhoneMobile;
                passenger.PhoneHome = p.PhoneHome;
                passenger.PhoneFax = p.PhoneFax;
                passenger.PhoneBusiness = p.PhoneBusiness;
                passenger.ReceivedFrom = p.ReceivedFrom;

                passenger.CreateBy = p.CreateBy;
                passenger.CreateDateTime = p.CreateDateTime;
                passenger.UpdateBy = p.UpdateBy;
                passenger.UpdateDateTime = p.UpdateDateTime;

                passenger.PassengerId = p.PassengerId;
                passenger.BookingId = p.BookingId;
                passenger.PassengerProfileId = p.PassengerProfileId;
                passenger.EmployeeNumber = p.EmployeeNumber;
                passenger.PassengerRoleRcd = p.PassengerRoleRcd;
                passenger.MemberLevelRcd = p.MemberLevelRcd;
                passenger.MemberNumber = p.MemberNumber;
                passenger.RedressNumber = p.RedressNumber;
                passenger.PnrName = p.PnrName;
                passenger.WheelchairFlag = p.WheelchairFlag;
                passenger.VipFlag = p.VipFlag;
                passenger.WindowSeatFlag = p.WindowSeatFlag;
                passenger.GuardianPassengerId = p.GuardianPassengerId;
                passenger.MemberAirlineRcd = p.MemberAirlineRcd;
            }

            return passenger;
        }
        public static IList<Mapping> ToListEntity(this  IList<Message.Booking.Mapping> m)
        {
            IList<Mapping> mappingList = null;

            if (m != null)
            {
                mappingList = new List<Mapping>();

                for (int i = 0; i < m.Count; i++)
                {
                    mappingList.Add(m[i].ToEntity());
                }
            }

            return mappingList;
        }
        public static Mapping ToEntity(this  Message.Booking.Mapping mapping)
        {
            Mapping passengerMapping = null;
            if (mapping != null)
            {
                passengerMapping = new Mapping();

                passengerMapping.OriginRcd = mapping.OriginRcd;
                passengerMapping.DestinationRcd = mapping.DestinationRcd;
                passengerMapping.DisplayName = mapping.DisplayName;
                passengerMapping.CreateBy = mapping.CreateBy;
                passengerMapping.UpdateBy = mapping.UpdateBy;
                passengerMapping.CreateDateTime = mapping.CreateDateTime;
                passengerMapping.UpdateDateTime = mapping.UpdateDateTime;
                #region Origin Flag
                passengerMapping.CurrencyRcd = mapping.CurrencyRcd;
                passengerMapping.RoutesTot = mapping.RoutesTot;
                passengerMapping.RoutesAvl = mapping.RoutesAvl;
                passengerMapping.RoutesB2c = mapping.RoutesB2c;
                passengerMapping.RoutesB2b = mapping.RoutesB2b;
                passengerMapping.RoutesB2s = mapping.RoutesB2s;
                passengerMapping.RoutesApi = mapping.RoutesApi;
                passengerMapping.RoutesB2t = mapping.RoutesB2t;
                #endregion
                #region Destination Flag
                passengerMapping.SegmentChangeFeeFlag = mapping.SegmentChangeFeeFlag;
                passengerMapping.TransitFlag = mapping.TransitFlag;
                passengerMapping.DirectFlag = mapping.DirectFlag;
                passengerMapping.AvlFlag = mapping.AvlFlag;
                passengerMapping.B2cFlag = mapping.B2cFlag;
                passengerMapping.B2bFlag = mapping.B2bFlag;
                passengerMapping.B2tFlag = mapping.B2tFlag;
                passengerMapping.DayRange = mapping.DayRange;
                passengerMapping.ShowRedressNumberFlag = mapping.ShowRedressNumberFlag;
                passengerMapping.RequirePassengerTitleFlag = mapping.RequirePassengerTitleFlag;
                passengerMapping.RequirePassengerGenderFlag = mapping.RequirePassengerGenderFlag;
                passengerMapping.RequireDateOfBirthFlag = mapping.RequireDateOfBirthFlag;
                passengerMapping.RequireDocumentDetailsFlag = mapping.RequireDocumentDetailsFlag;
                passengerMapping.RequirePassengerWeightFlag = mapping.RequirePassengerWeightFlag;
                passengerMapping.SpecialServiceFeeFlag = mapping.SpecialServiceFeeFlag;
                passengerMapping.ShowInsuranceOnWebFlag = mapping.ShowInsuranceOnWebFlag;
                #endregion

                #region "Property"
                passengerMapping.FlightId = mapping.FlightId;
                passengerMapping.ExchangedSegmentId = mapping.ExchangedSegmentId;
                passengerMapping.AirlineRcd = mapping.AirlineRcd;
                passengerMapping.FlightNumber = mapping.FlightNumber;
                passengerMapping.RefundableFlag = mapping.RefundableFlag;
                passengerMapping.GroupFlag = mapping.GroupFlag;
                passengerMapping.NonRevenueFlag = mapping.NonRevenueFlag;
                passengerMapping.EticketFlag = mapping.EticketFlag;
                passengerMapping.FareReduction = mapping.FareReduction;
                passengerMapping.DepartureDate = mapping.DepartureDate;
                passengerMapping.ArrivalDate = mapping.ArrivalDate;
                passengerMapping.OdOriginRcd = mapping.OdOriginRcd;
                passengerMapping.OdDestinationRcd = mapping.OdDestinationRcd;
                passengerMapping.WaitlistFlag = mapping.WaitlistFlag;
                passengerMapping.OverbookFlag = mapping.OverbookFlag;
                passengerMapping.FlightConnectionId = mapping.FlightConnectionId;
                passengerMapping.AdvancedPurchaseFlag = mapping.AdvancedPurchaseFlag;
                passengerMapping.JourneyTime = mapping.JourneyTime;
                passengerMapping.DepartureTime = mapping.DepartureTime;
                passengerMapping.ArrivalTime = mapping.ArrivalTime;


                passengerMapping.TransitFlightId = mapping.TransitFlightId;
                passengerMapping.TransitFareId = mapping.TransitFareId;
                passengerMapping.TransitDepartureDate = mapping.TransitDepartureDate;
                passengerMapping.TransitArrivalDate = mapping.TransitArrivalDate;
                passengerMapping.FareId = mapping.FareId;
                passengerMapping.BookingClassRcd = mapping.BookingClassRcd;
                passengerMapping.BoardingClassRcd = mapping.BoardingClassRcd;

                passengerMapping.PlannedDepartureTime = mapping.PlannedDepartureTime;
                passengerMapping.PlannedArrivalTime = mapping.PlannedArrivalTime;
                passengerMapping.DepartureDayofweek = mapping.DepartureDayofweek;
                passengerMapping.ArrivalDayofweek = mapping.ArrivalDayofweek;
                passengerMapping.DepartureMonth = mapping.DepartureMonth;
                #endregion

                #region "Property"

                passengerMapping.BookingSegmentId = mapping.BookingSegmentId;
                passengerMapping.BookingId = mapping.BookingId;
                passengerMapping.NumberOfUnits = mapping.NumberOfUnits;
                passengerMapping.InfoSegmentFlag = mapping.InfoSegmentFlag;
                passengerMapping.HighPriorityWaitlistFlag = mapping.HighPriorityWaitlistFlag;
                passengerMapping.SeatmapFlag = mapping.SeatmapFlag;
                passengerMapping.TempSeatmapFlag = mapping.TempSeatmapFlag;
                passengerMapping.AllowWebCheckinFlag = mapping.AllowWebCheckinFlag;
                passengerMapping.CloseWebSalesFlag = mapping.CloseWebSalesFlag;
                passengerMapping.ExcludeQuoteFlag = mapping.ExcludeQuoteFlag;
                passengerMapping.CurrencyRate = mapping.CurrencyRate;
                passengerMapping.OpenSequence = mapping.OpenSequence;
                passengerMapping.NumberOfStops = mapping.NumberOfStops;
                #endregion


                #region Passenger information
                passengerMapping.PassengerId = mapping.PassengerId;
                passengerMapping.Lastname = mapping.Lastname;
                passengerMapping.Firstname = mapping.Firstname;
                passengerMapping.GenderTypeRcd = mapping.GenderTypeRcd;
                passengerMapping.TitleRcd = mapping.TitleRcd;
                passengerMapping.PassengerTypeRcd = mapping.PassengerTypeRcd;
                passengerMapping.DateOfBirth = mapping.DateOfBirth;
                passengerMapping.PassengerStatusRcd = mapping.PassengerStatusRcd;
                passengerMapping.PassengerCheckinStatusRcd = mapping.PassengerCheckinStatusRcd;

                passengerMapping.Middlename = mapping.Middlename;
                passengerMapping.DocumentTypeRcd = mapping.DocumentTypeRcd;
                passengerMapping.PassportNumber = mapping.PassportNumber;
                passengerMapping.PassportIssuePlace = mapping.PassportIssuePlace;
                passengerMapping.PassportIssueDate = mapping.PassportIssueDate;
                passengerMapping.PassportExpiryDate = mapping.PassportExpiryDate;
                passengerMapping.Sendmail = mapping.Sendmail;
                passengerMapping.ClientNumber = mapping.ClientNumber;
                passengerMapping.PassengerProfileId = mapping.PassengerProfileId;
                passengerMapping.ClientProfileId = mapping.ClientProfileId;
                passengerMapping.EmployeeNumber = mapping.EmployeeNumber;
                passengerMapping.NationalityRcd = mapping.NationalityRcd;
                passengerMapping.ContactEmail = mapping.ContactEmail;

                #endregion
                #region Ticket information
                passengerMapping.InventoryClassRcd = mapping.InventoryClassRcd;
                passengerMapping.SeatNumber = mapping.SeatNumber;
                passengerMapping.SeatRow = mapping.SeatRow;
                passengerMapping.SeatColumn = mapping.SeatColumn;
                passengerMapping.NetTotal = mapping.NetTotal;
                passengerMapping.TaxAmount = mapping.TaxAmount;
                passengerMapping.FareAmount = mapping.FareAmount;
                passengerMapping.YqAmount = mapping.YqAmount;
                passengerMapping.BaseTicketingFeeAmount = mapping.BaseTicketingFeeAmount;
                passengerMapping.TicketingFeeAmount = mapping.TicketingFeeAmount;
                passengerMapping.ReservationFeeAmount = mapping.ReservationFeeAmount;
                passengerMapping.CommissionAmount = mapping.CommissionAmount;
                passengerMapping.FareVat = mapping.FareVat;
                passengerMapping.TaxVat = mapping.TaxVat;
                passengerMapping.YqVat = mapping.YqVat;
                passengerMapping.TicketingFeeVat = mapping.TicketingFeeVat;
                passengerMapping.ReservationFeeVat = mapping.ReservationFeeVat;
                passengerMapping.FareAmountIncl = mapping.FareAmountIncl;
                passengerMapping.TaxAmountIncl = mapping.TaxAmountIncl;
                passengerMapping.YqAmountIncl = mapping.YqAmountIncl;

                passengerMapping.PublicFareAmountIncl = mapping.PublicFareAmountIncl;
                passengerMapping.PublicFareAmount = mapping.PublicFareAmount;

                passengerMapping.CommissionAmountIncl = mapping.CommissionAmountIncl;
                passengerMapping.CommissionPercentage = mapping.CommissionPercentage;
                passengerMapping.TicketingFeeAmountIncl = mapping.TicketingFeeAmountIncl;
                passengerMapping.ReservationFeeAmountIncl = mapping.ReservationFeeAmountIncl;
                passengerMapping.AcctNetTotal = mapping.AcctNetTotal;
                passengerMapping.AcctTaxAmount = mapping.AcctTaxAmount;
                passengerMapping.AcctFareAmount = mapping.AcctFareAmount;
                passengerMapping.AcctYqAmount = mapping.AcctYqAmount;
                passengerMapping.AcctTicketingFeeAmount = mapping.AcctTicketingFeeAmount;
                passengerMapping.AcctReservationFeeAmount = mapping.AcctReservationFeeAmount;
                passengerMapping.AcctCommissionAmount = mapping.AcctCommissionAmount;
                passengerMapping.AcctFareAmountIncl = mapping.AcctFareAmountIncl;
                passengerMapping.AcctTaxAmountIncl = mapping.AcctTaxAmountIncl;
                passengerMapping.AcctYqAmountIncl = mapping.AcctYqAmountIncl;
                passengerMapping.AcctCommissionAmountIncl = mapping.AcctCommissionAmountIncl;
                passengerMapping.AcctTicketingFeeAmountIncl = mapping.AcctTicketingFeeAmountIncl;
                passengerMapping.AcctReservationFeeAmountIncl = mapping.AcctReservationFeeAmountIncl;
                passengerMapping.FareCode = mapping.FareCode;
                passengerMapping.FareDateTime = mapping.FareDateTime;
                passengerMapping.ETicketFlag = mapping.ETicketFlag;
                passengerMapping.StandbyFlag = mapping.StandbyFlag;
                passengerMapping.TransferableFareFlag = mapping.TransferableFareFlag;
                passengerMapping.AgencyCode = mapping.AgencyCode;
                passengerMapping.TicketNumber = mapping.TicketNumber;
                passengerMapping.TicketingDateTime = mapping.TicketingDateTime;
                passengerMapping.TicketingBy = mapping.TicketingBy;
                passengerMapping.CheckInSequence = mapping.CheckInSequence;
                passengerMapping.GroupSequence = mapping.GroupSequence;
                passengerMapping.UnloadBy = mapping.UnloadBy;
                passengerMapping.UnloadDateTime = mapping.UnloadDateTime;
                passengerMapping.NumberOfPieces = mapping.NumberOfPieces;
                passengerMapping.BaggageWeight = mapping.BaggageWeight;
                passengerMapping.ExcessBaggageWeight = mapping.ExcessBaggageWeight;
                passengerMapping.CheckInBaggageWeight = mapping.CheckInBaggageWeight;
                passengerMapping.CheckInBy = mapping.CheckInBy;
                passengerMapping.CancelBy = mapping.CancelBy;
                passengerMapping.ExchangedDateTime = mapping.ExchangedDateTime;
                passengerMapping.CancelDateTime = mapping.CancelDateTime;
                passengerMapping.RestrictionText = mapping.RestrictionText;
                passengerMapping.EndorsementText = mapping.EndorsementText;
                passengerMapping.ExcludePricingFlag = mapping.ExcludePricingFlag;
                passengerMapping.CreateName = mapping.CreateName;
                passengerMapping.UpdateName = mapping.UpdateName;
                passengerMapping.VoidDateTime = mapping.VoidDateTime;
                passengerMapping.RefundCharge = mapping.RefundCharge;
                passengerMapping.RefundableAmount = mapping.RefundableAmount;
                passengerMapping.RefundDateTime = mapping.RefundDateTime;
                passengerMapping.PaymentAmount = mapping.PaymentAmount;
                passengerMapping.MappingSort = mapping.MappingSort;

                passengerMapping.CheckInDateTime = mapping.CheckInDateTime;
                passengerMapping.OnwardDepartureDate = mapping.OnwardDepartureDate;
                passengerMapping.ETicketStatus = mapping.ETicketStatus;
                passengerMapping.HandLuggageFlag = mapping.HandLuggageFlag;
                passengerMapping.HandNumberOfPieces = mapping.HandNumberOfPieces;
                passengerMapping.HandBaggageWeight = mapping.HandBaggageWeight;
                passengerMapping.FreeSeatingFlag = mapping.FreeSeatingFlag;
                passengerMapping.FareTypeRcd = mapping.FareTypeRcd;
                passengerMapping.RedemptionPoints = mapping.RedemptionPoints;
                passengerMapping.SeatFeeRcd = mapping.SeatFeeRcd;
                passengerMapping.RefundWithChargeHours = mapping.RefundWithChargeHours;
                passengerMapping.RefundNotPossibleHours = mapping.RefundNotPossibleHours;
                passengerMapping.DutyTravelFlag = mapping.DutyTravelFlag;
                passengerMapping.NotValidAfterDate = mapping.NotValidAfterDate;
                passengerMapping.NotValidBeforeDate = mapping.NotValidBeforeDate;
                passengerMapping.AdvancedSeatingFlag = mapping.AdvancedSeatingFlag;
                passengerMapping.FareColumn = mapping.FareColumn;
                #endregion
                #region Ignore XML Section
                passengerMapping.PieceAllowance = mapping.PieceAllowance;
                passengerMapping.BoardingTime = mapping.BoardingTime;
                passengerMapping.ItFareFlag = mapping.ItFareFlag;
                #endregion
            }

            return passengerMapping;
        }
        public static IList<Payment> ToListEntity(this  IList<Message.Booking.Payment> pay)
        {
            IList<Payment> paymentList = null;

            if (pay != null)
            {
                paymentList = new List<Payment>();
                for (int i = 0; i < pay.Count; i++)
                {
                    paymentList.Add(pay[i].ToEntity());
                }
            }
            return paymentList;
        }
        public static Payment ToEntity(this  Message.Booking.Payment pay)
        {
            Payment payment = null;
            if (pay != null)
            {
                payment = new Payment();
                payment.BookingPaymentId = pay.BookingPaymentId;

                if (pay.BookingPaymentId != Guid.Empty)
                    payment.BookingPaymentId = pay.BookingPaymentId;
                else
                    payment.BookingPaymentId = Guid.NewGuid();

                payment.BookingSegmentId = pay.BookingSegmentId;
                payment.BookingId = pay.BookingId;
                payment.VoucherPaymentId = pay.VoucherPaymentId;
                payment.FormOfPaymentRcd = pay.FormOfPaymentRcd;
                payment.CurrencyRcd = pay.CurrencyRcd;
                payment.ReceiveCurrencyRcd = pay.ReceiveCurrencyRcd;
                payment.AgencyPaymentTypeRcd = pay.AgencyPaymentTypeRcd;
                payment.AgencyCode = pay.AgencyCode;
                payment.DebitAgencyCode = pay.DebitAgencyCode;
                payment.PaymentAmount = pay.PaymentAmount;
                payment.ReceivePaymentAmount = pay.ReceivePaymentAmount;
                payment.AcctPaymentAmount = pay.AcctPaymentAmount;
                payment.PaymentBy = pay.PaymentBy;
                payment.PaymentDateTime = pay.PaymentDateTime;
                payment.PaymentDueDateTime = pay.PaymentDueDateTime;
                payment.DocumentAmount = pay.DocumentAmount;
                payment.VoidBy = pay.VoidBy;
                payment.ExpiryMonth = pay.ExpiryMonth;
                payment.ExpiryYear = pay.ExpiryYear;
                payment.VoidDateTime = pay.VoidDateTime;
                payment.RecordLocator = pay.RecordLocator;
                payment.CvvCode = pay.CvvCode;
                payment.NameOnCard = pay.NameOnCard;
                payment.DocumentNumber = pay.DocumentNumber;
                payment.FormOfPaymentSubtypeRcd = pay.FormOfPaymentSubtypeRcd;
                payment.City = pay.City;
                payment.State = pay.State;
                payment.Street = pay.Street;
                payment.PoBox = pay.PoBox;
                payment.AddressLine1 = pay.AddressLine1;
                payment.AddressLine2 = pay.AddressLine2;
                payment.District = pay.District;
                payment.Province = pay.Province;
                payment.ZipCode = pay.ZipCode;
                payment.CountryRcd = pay.CountryRcd;
                payment.CreateBy = pay.CreateBy;
                payment.CreateDateTime = pay.CreateDateTime;
                payment.UpdateBy = pay.UpdateBy;
                payment.UpdateDateTime = pay.UpdateDateTime;
                payment.PosIndicator = pay.PosIndicator;
                payment.IssueMonth = pay.IssueMonth;
                payment.IssueYear = pay.IssueYear;
                payment.IssueNumber = pay.IssueNumber;
                payment.ClientProfileId = pay.ClientProfileId;
                payment.TransactionReference = pay.TransactionReference;
                payment.ApprovalCode = pay.ApprovalCode;
                payment.BankName = pay.BankName;
                payment.BankCode = pay.BankCode;
                payment.BankIban = pay.BankIban;
                payment.IpAddress = pay.IpAddress;
                payment.PaymentReference = pay.PaymentReference;
                payment.AllocatedAmount = pay.AllocatedAmount;
                payment.PaymentTypeRcd = pay.PaymentTypeRcd;
                payment.RefundedAmount = pay.RefundedAmount;

                payment.PaymentNumber = pay.PaymentNumber;
                payment.PaymentRemark = pay.PaymentRemark;


            }

            return payment;
        }
        public static IList<Fee> ToListEntity(this  IList<Message.Booking.Fee> f)
        {
            IList<Fee> feeList = null;

            if (f != null)
            {
                feeList = new List<Fee>();
                for (int i = 0; i < f.Count; i++)
                {
                    feeList.Add(f[i].ToEntity());
                }
            }

            return feeList;
        }
        public static Fee ToEntity(this  Message.Booking.Fee f)
        {
            Fee fee = null;
            if (f != null)
            {
                fee = new Fee();

                if (f.BookingFeeId != Guid.Empty)
                    fee.BookingFeeId = f.BookingFeeId;
                else
                    fee.BookingFeeId = Guid.NewGuid();

                fee.FeeId = f.FeeId;
                fee.FeeAmount = f.FeeAmount;
                fee.BookingId = f.BookingId;
                fee.PassengerId = f.PassengerId;
                fee.CurrencyRcd = f.CurrencyRcd;
                fee.AcctFeeAmount = f.AcctFeeAmount;
                fee.VatPercentage = f.VatPercentage;
                fee.FeeAmountIncl = f.FeeAmount;
                fee.AcctFeeAmountIncl = f.AcctFeeAmountIncl;
                fee.FeeRcd = f.FeeRcd;
                fee.DisplayName = f.DisplayName;
                fee.AccountFeeBy = f.AccountFeeBy;
                fee.AccountFeeDateTime = f.AccountFeeDateTime;
                fee.VoidDateTime = f.VoidDateTime;
                fee.VoidBy = f.VoidBy;
                fee.PaymentAmount = f.PaymentAmount;
                fee.CreateBy = f.CreateBy;
                fee.CreateDateTime = f.CreateDateTime;
                fee.UpdateBy = f.UpdateBy;
                fee.UpdateDateTime = f.UpdateDateTime;
                fee.BookingSegmentId = f.BookingSegmentId;
                fee.AgencyCode = f.AgencyCode;
                fee.PassengerSegmentServiceId = f.PassengerSegmentServiceId;

                if (!string.IsNullOrEmpty(f.FeeCategoryRcd))
                {
                    if (f.FeeCategoryRcd == "BAGSALES")
                        fee.FeeCategoryRcd = "BAGF";
                    else
                        fee.FeeCategoryRcd = f.FeeCategoryRcd;
                }
                else
                {
                    fee.FeeCategoryRcd = "OTH";
                }

                fee.OriginRcd = f.OriginRcd;
                fee.DestinationRcd = f.DestinationRcd;
                fee.OdOriginRcd = f.OdOriginRcd;
                fee.OdDestinationRcd = f.OdDestinationRcd;
                fee.NumberOfUnits = f.NumberOfUnits;
                fee.TotalAmount = f.TotalAmount;
                fee.TotalAmountIncl = f.TotalAmountIncl;
                fee.ManualFeeFlag = f.ManualFeeFlag;
                fee.OdFlag = f.OdFlag;
                fee.SkipFareAllowanceFlag = f.SkipFareAllowanceFlag;
                fee.FeeLevel = f.FeeLevel;
                fee.FeeCalculationRcd = f.FeeCalculationRcd;
                fee.MinimumFeeAmountFlag = f.MinimumFeeAmountFlag;
                fee.FeePercentage = f.FeePercentage;
                fee.ChangeComment = f.ChangeComment;
                fee.Comment = f.Comment;
                fee.TotalFeeAmount = f.TotalFeeAmount;
                fee.TotalFeeAmountIncl = f.TotalFeeAmountIncl;
                fee.Units = f.Units;
                fee.ChargeCurrencyRcd = f.ChargeCurrencyRcd;
                fee.ChargeAmount = f.ChargeAmount;
                fee.ChargeAmountIncl = f.ChargeAmountIncl;
                fee.DocumentNumber = f.DocumentNumber;
                fee.DocumentDateTime = f.DocumentDateTime;
                fee.ExternalReference = f.ExternalReference;
                fee.VendorRcd = f.VendorRcd;
                fee.BaggageFeeOptionId = f.BaggageFeeOptionId;
                fee.NewRecord = f.NewRecord;
                fee.SelectedFee = f.SelectedFee;
                fee.WeightLbs = f.WeightLbs;
                fee.WeightKgs = f.WeightKgs;
                fee.MpdNumber = f.MpdNumber;
                fee.UnitRcd = f.UnitRcd;
            }

            return fee;
        }
        public static IList<Remark> ToListEntity(this  IList<Message.Booking.Remark> r)
        {
            IList<Remark> remarkList = null;

            if (r != null)
            {
                remarkList = new List<Remark>();
                for (int i = 0; i < r.Count; i++)
                {
                    remarkList.Add(r[i].ToEntity());
                }
            }

            return remarkList;
        }
        public static Remark ToEntity(this  Message.Booking.Remark r)
        {
            Remark remark = null;
            if (r != null)
            {
                remark = new Remark();

                if (r.BookingRemarkId != Guid.Empty)
                    remark.BookingRemarkId = r.BookingRemarkId;
                else
                    remark.BookingRemarkId = Guid.NewGuid();

                remark.RemarkTypeRcd = r.RemarkTypeRcd;
                remark.TimelimitDateTime = r.TimelimitDateTime;
                remark.CompleteFlag = r.CompleteFlag;
                remark.RemarkText = r.RemarkText;
                remark.BookingId = r.BookingId;
                remark.AddedBy = r.AddedBy;
                remark.ClientProfileId = r.ClientProfileId;
                remark.AgencyCode = r.AgencyCode;
                remark.CreateBy = r.CreateBy;
                remark.CreateDateTime = r.CreateDateTime;
                remark.UpdateBy = r.UpdateBy;
                remark.UpdateDateTime = r.UpdateDateTime;
                remark.SystemFlag = r.SystemFlag;
                remark.UtcTimelimitDateTime = r.UtcTimelimitDateTime;
                remark.Nickname = r.Nickname;
                remark.ProtectedFlag = r.ProtectedFlag;
                remark.WarningFlag = r.WarningFlag;
                remark.ProcessMessageFlag = r.ProcessMessageFlag; 

            }

            return remark;
        }
        public static IList<PassengerService> ToListEntity(this  IList<Message.Booking.PassengerService> service)
        {
            IList<PassengerService> serviceList = null;

            if (service != null)
            {
                serviceList = new List<PassengerService>();
                for (int i = 0; i < service.Count; i++)
                {
                    serviceList.Add(service[i].ToEntity());
                }
            }

            return serviceList;
        }
        public static PassengerService ToEntity(this  Message.Booking.PassengerService s)
        {
            PassengerService service = null;
            if (s != null)
            {
                service = new PassengerService();
                if (s.PassengerSegmentServiceId != Guid.Empty)
                    service.PassengerSegmentServiceId = s.PassengerSegmentServiceId;
                else
                    service.PassengerSegmentServiceId = Guid.NewGuid();

                service.PassengerId = s.PassengerId;
                service.BookingSegmentId = s.BookingSegmentId;
                service.SpecialServiceStatusRcd = s.SpecialServiceStatusRcd;
                service.SpecialServiceChangeStatusRcd = s.SpecialServiceChangeStatusRcd;
                service.SpecialServiceRcd = s.SpecialServiceRcd;
                service.ServiceText = s.ServiceText;
                service.CreateBy = s.CreateBy;
                service.CreateDateTime = s.CreateDateTime;
                service.UpdateBy = s.UpdateBy;
                service.UpdateDateTime = s.UpdateDateTime;
                service.FlightId = s.FlightId;
                service.FeeId = s.FeeId;
                service.NumberOfUnits = s.NumberOfUnits;
                service.OriginRcd = s.OriginRcd;
                service.DestinationRcd = s.DestinationRcd;
                service.NewRecord = s.NewRecord;
                service.DisplayName = s.DisplayName;
                service.CutOffTime = s.CutOffTime;
                service.StatusCode = s.StatusCode;
                service.HelpText = s.HelpText;
                service.SpecialServiceGroupRcd = s.SpecialServiceGroupRcd;
                service.ServiceOnRequestFlag = s.ServiceOnRequestFlag;
                service.TextAllowedFlag = s.TextAllowedFlag;
                service.ManifestFlag = s.ManifestFlag;
                service.TextRequiredFlag = s.TextRequiredFlag;
                service.IncludePassengerNameFlag = s.IncludePassengerNameFlag;
                service.IncludeFlightSegmentFlag = s.IncludeFlightSegmentFlag;
                service.IncludeActionCodeFlag = s.IncludeActionCodeFlag;
                service.IncludeNumberOfServiceFlag = s.IncludeNumberOfServiceFlag;
                service.IncludeCateringFlag = s.IncludeCateringFlag;
                service.IncludePassengerAssistanceFlag = s.IncludePassengerAssistanceFlag;
                service.ServiceSupportedFlag = s.ServiceSupportedFlag;
                service.SendInterlineReplyFlag = s.SendInterlineReplyFlag;
                service.InventoryControlFlag = s.InventoryControlFlag;

                service.UnitRcd = s.UnitRcd;
            }

            return service;
        }
        public static IList<Tax> ToListEntity(this  IList<Message.Booking.Tax> tax)
        {
            IList<Tax> taxList = null;

            if (tax != null)
            {
                taxList = new List<Tax>();
                for (int i = 0; i < tax.Count; i++)
                {
                    taxList.Add(tax[i].ToEntity());
                }
            }

            return taxList;
        }
        public static Tax ToEntity(this  Message.Booking.Tax t)
        {
            Tax tax = null;

            if (t != null)
            {
                tax = new Tax();

                tax.PassengerSegmentTaxId = t.PassengerSegmentTaxId;
                tax.TaxAmount = t.TaxAmount;
                tax.TaxAmountIncl = t.TaxAmountIncl;
                tax.AcctAmount = t.AcctAmount;
                tax.AcctAmountIncl = t.AcctAmountIncl;
                tax.SalesAmount = t.SalesAmount;
                tax.SalesAmountIncl = t.SalesAmountIncl;
                tax.TaxRcd = t.TaxRcd;
                tax.PassengerId = t.PassengerId;
                tax.TaxId = t.TaxId;
                tax.BookingSegmentId = t.BookingSegmentId;
                tax.TaxCurrencyRcd = t.TaxCurrencyRcd;
                tax.SalesCurrencyRcd = t.SalesCurrencyRcd;
                tax.DisplayName = t.DisplayName;
                tax.SummarizeUp = t.SummarizeUp;
                tax.CoverageType = t.CoverageType;
                tax.CreateBy = t.CreateBy;
                tax.UpdateBy = t.UpdateBy;
                tax.CreateDateTime = t.CreateDateTime;
                tax.UpdateDateTime = t.UpdateDateTime;
                tax.VatPercentage = t.VatPercentage; 

            }

            return tax;
        }

    }
}
