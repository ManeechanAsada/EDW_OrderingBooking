using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using Avantik.Web.Service.Entity.Flight;
using Avantik.Web.Service.Infrastructure;

namespace Avantik.Web.Service.Repository.Contract.Flight
{
    public interface IAvailabilityRepository
    {
        IEnumerable<Availability> FindAvailability(string originRcd,
                                                   string destinationRcd,
                                                   string bookingClass,
                                                   string boardingClass,
                                                   string otherPax,
                                                   string agencyCode,
                                                   string currencyCode,
                                                   string dayTime,
                                                   string odOriginRcd,
                                                   string odDestinationRcd,
                                                   string promotionCode,
                                                   AvailabilityFareTypes fareType,
                                                   string languageCode,
                                                   string ipAddress,
                                                   DateTime dateFrom,
                                                   DateTime dateTo,
                                                   DateTime bookingDate,
                                                   Guid flightId,
                                                   Guid fareId,
                                                   Int16 adult,
                                                   Int16 child,
                                                   Int16 infant,
                                                   Int16 other,
                                                   Int16 nonstop,
                                                   Int16 groupFare,
                                                   Int16 refundable,
                                                   Int16 cancelled,
                                                   Int16 departed,
                                                   Int16 waitlist,
                                                   Int16 soldout,
                                                   Int16 itFares,
                                                   Int16 staffFares,
                                                   Int16 includeFares,
                                                   decimal maxFare,
                                                   bool mapWithFare,
                                                   string transactionReference);
    }
}
