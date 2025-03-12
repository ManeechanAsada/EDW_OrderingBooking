using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Avantik.Web.Service.Entity.Flight;
using Avantik.Web.Service.Infrastructure
;
namespace Avantik.Web.Service.Model.Contract
{
    public interface IAvailabilityService
    {
        Availabilities GetAvailability(string otherPassengerType,
                                       string boardingClass,
                                       string bookingClass,
                                       string dayTimeIndicator,
                                       string returnDayTimeIndicator,
                                       string agencyCode,
                                       string currencyCode,
                                       string transitPoint,
                                       string promotionCode,
                                       AvailabilityFareTypes fareType,
                                       string languageCode,
                                       string ipAddress,
                                       string originRcd,
                                       string destinationRcd,
                                       string odOriginRcd,
                                       string odDestinationRcd,
                                       DateTime fromDate,
                                       DateTime toDate,
                                       DateTime returnFromDate,
                                       DateTime returnToDate,
                                       DateTime bookingDate,
                                       decimal maxAmount,
                                       byte adult,
                                       byte child,
                                       byte infant,
                                       byte other,
                                       Guid flightId,
                                       Guid fareId,
                                       short nonStopOnly,
                                       short includeDeparted,
                                       short includeCancelled,
                                       short includeWaitlisted,
                                       short includeSoldOut,
                                       short includeFares,
                                       short refundable,
                                       short returnRefundable,
                                       short groupFares,
                                       short iTFaresOnly,
                                       short staffFares,
                                       bool applyFareLogic,
                                       bool showClose,
                                       short unknownTransit,
                                       bool mapWithFares,
                                       bool noVat,
                                       string transactionReference);
    }
}
