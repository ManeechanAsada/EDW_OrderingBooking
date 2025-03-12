using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Avantik.Web.Service.Entity.Flight;
using Avantik.Web.Service.Infrastructure;
using Avantik.Web.Service.Entity.Booking;
using Avantik.Web.Service.Entity;

namespace Avantik.Web.Service.Model.Contract
{
    public interface IFeeService
    {
        IList<Entity.Fee> GetFee(string strFeeRcd,
                          string strCurrencyCode,
                          string strAgencyCode,
                          string strClass,
                          string strFareBasis,
                          string strOrigin,
                          string strDestination,
                          string strFlightNumber,
                          DateTime dtDate,
                          string strLanguage,
                          bool bNovat);

        IList<Entity.Booking.Fee> CalculateFeesBookingCreate(string AgencyCode,
                                                        string currency,
                                                        Booking booking,
                                                        string strLanguage);
        IList<Entity.Booking.Fee> CalculateFeesBookingChange(string AgencyCode,
                                            string currency,
                                            Booking booking,
                                            string strLanguage);

        IList<Entity.Booking.Fee> CalculateFeesSeatAssignment(string AgencyCode,
                                               string currency,
                                               Booking booking,
                                               string strLanguage,
                                               bool bNovat);
        IList<Entity.Booking.Fee> CalculateFeesNameChange(string AgencyCode,
                                            string currency,
                                            Booking booking,
                                            string strLanguage, string strUserId);
        Entity.Booking.Booking CalculateFeesSpecialService(string AgencyCode,
                                                string currency,
                                                Booking booking,
                                                string strLanguage,
                                                bool bNovat);
        List<ServiceFee> GetSegmentFee(string agencyCode,
                               string currencyCode,
                               string languageCode,
                               int numberOfPassenger,
                               int numberOfInfant,
                               IList<PassengerService> services,
                               IList<SegmentService> segmentService,
                               bool SpecialService,
                               bool bNovat);

        IList<Entity.Booking.Fee> GetBaggageFee(IList<Mapping> mapping, Guid bookingSegmentId, Guid passengerId, string agencyCode, string languageCode, int maxUnits, IList<Entity.Booking.Fee> fees, bool bNovat);
    }
}
