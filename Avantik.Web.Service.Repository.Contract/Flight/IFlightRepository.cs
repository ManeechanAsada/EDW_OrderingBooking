using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using Avantik.Web.Service.Entity.Flight;

namespace Avantik.Web.Service.Repository.Contract.Flight
{
    public interface IFlightRepository
    {
        IEnumerable<AvailabilityRoute> GetFlightAvailabilityRoute(string originRcd, string destinationRcd);
        IEnumerable<AvailabilityParing> GetFlightAvailabilityParing(string originRcd,
                                                                    string destinationRcd,
                                                                    string transit,
                                                                    DateTime dateFrom,
                                                                    DateTime dateTo);
        IEnumerable<AvailabilityQuoteTax> GetFlightAvailabilityQuoteTax(string originRcd,
                                                                        string destinationRcd,
                                                                        string currencyRcd,
                                                                        string passengerTypeRcd,
                                                                        string agencyCode,
                                                                        DateTime flightDate,
                                                                        DateTime bookingDate);
    }
}
