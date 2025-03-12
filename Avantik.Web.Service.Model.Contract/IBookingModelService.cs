using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Avantik.Web.Service.Entity.Flight;
using Avantik.Web.Service.Infrastructure;
using Avantik.Web.Service.Entity.Booking;
using Avantik.Web.Service.Entity.Agency;

namespace Avantik.Web.Service.Model.Contract
{
    public interface IBookingModelService 
    {
          bool SaveBooking(Booking booking,
                           bool createTickets,
                           bool readBooking,
                           bool readOnly,
                           bool bSetLock,
                           bool bCheckSeatAssignment,
                           bool bCheckSessionTimeOut);
          Booking ReadBooking(string bookingId,
                                     string bookingReference,
                                     double bookingNumber,
                                     bool bReadonly,
                                     bool bSeatLock,
                                     string userId,
                                     bool bReadHeader,
                                     bool bReadSegment,
                                     bool bReadPassenger,
                                     bool bReadRemark,
                                     bool bReadPayment,
                                     bool bReadMapping,
                                     bool bReadService,
                                     bool bReadTax,
                                     bool bReadFee,
                                     bool bReadOd,
                                     string releaseBookingId,
                                     string CompleteRemarkId);
          bool CancelBooking(string bookingId,
                                   string bookingReference,
                                   double bookingNumber,
                                   string userId,
                                   string agencyCode,
                                   bool bWaveFee,
                                   bool bVoid);
          bool CancelSegment(string bookingId,
                                     string segmentId,
                                     string userId,
                                     string agencyCode,
                                     bool bWaveFee,
                                     bool bVoid);
          Booking BookFlight(string agencyCode,
                          string currency,
                          IList<Flight> flight,
                          string bookingId,
                          short adults,
                          short children,
                          short infants,
                          short others,
                          string strOthers,
                          string userId,
                          string strIpAddress,
                          string strLanguageCode,
                          bool bNoVat);
          Booking ModifyBooking(string strBookingId,
                                IList<Flight> flight,
                                IList<PassengerService> services,
                                IList<Entity.SeatAssign> seatAssign,
                                IList<Entity.Booking.Fee> baggageFee,
                                IList<Entity.NameChange> NameChange,
                                IList<Payment> payments,
                                string strUserId,
                                string agencyCode,
                                string currencyRcd,
                                string languageRcd,
                                bool bNoVat,
                                string actionCode);

          Booking OrderingBooking(string strBookingId,
                        IList<Entity.Booking.FlightSegment> segments,
                        IList<Entity.Booking.Mapping> mappings,
                        IList<Entity.Booking.Fee> fees,
                        IList<Entity.Booking.PassengerService> services,
                        IList<Entity.Booking.Tax> taxes,
                        string strUserId,
                        string agencyCode,
                        string currencyRcd,
                        string languageRcd);


          bool UpdatePassengerInfo(string strBookingId,
                        IList<Passenger> passenger,
                        IList<Mapping> mapping,
                        string strUserId,
                        string agencyCode,
                        bool bNoVat);
          bool UpdateContactDetail(string strBookingId,
                                        BookingHeader header,
                                        string strUserId,
                                        string agencyCode,
                                        bool bNoVat);
          bool UpdateTicket(string strBookingId,
                        IList<Mapping> mapping,
                        string strUserId,
                        string agencyCode,
                        bool bNoVat);
          Booking CalculateBookingChange(string strBookingId,
                                  IList<Flight> flight,
                                  string strUserId,
                                  string agencyCode,
                                  bool bNoVat);
          Booking GetQuoteSummary(IList<FlightSegment> flights, 
                                  IList<Passenger> passengers, 
                                  string agencyCode, 
                                  string language, 
                                  string currencyCode, 
                                  bool bNovat);
          Booking ChangeFlightFee(string strBookingId,
                          IList<Flight> flight,
                          string strUserId,
                          string agencyCode,
                          string currency,
                          string language,
                          bool bNoVat);

    }
}
