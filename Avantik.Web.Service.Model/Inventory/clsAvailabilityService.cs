using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using Avantik.Web.Service.Entity.Flight;
using Avantik.Web.Service.Repository.Contract.Flight;
using Avantik.Web.Service.Infrastructure;
using Avantik.Web.Service.Model.Contract;

namespace Avantik.Web.Service.Model
{
    public class AvailabilityService : IAvailabilityService
    {
        IAvailabilityRepository _availabilityRepository;
        IFlightRepository _flightRepository;

        AvailabilityTypes _availabilityType;

        public AvailabilityService(AvailabilityTypes availabilityType)
        {
            _availabilityType = availabilityType;
        }
        public AvailabilityService(AvailabilityTypes availabilityType,
                            IAvailabilityRepository availabilityRepository,
                            IFlightRepository flightRepository)
        {
            _availabilityType = availabilityType;
            _availabilityRepository = availabilityRepository;
            _flightRepository = flightRepository;
        }
        public Availabilities GetAvailability(string otherPassengerType,
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
                                               string transactionReference)
        {
            Availabilities availabilities = new Availabilities();

            //Find outbound availability
            availabilities.FlightAvailabilityOutbound = GetAvailableFlight(otherPassengerType,
                                                                           boardingClass,
                                                                           bookingClass,
                                                                           dayTimeIndicator,
                                                                           agencyCode,
                                                                           currencyCode,
                                                                           transitPoint,
                                                                           promotionCode,
                                                                           fareType,
                                                                           languageCode,
                                                                           ipAddress,
                                                                           originRcd,
                                                                           destinationRcd,
                                                                           odOriginRcd,
                                                                           odDestinationRcd,
                                                                           fromDate,
                                                                           toDate,
                                                                           bookingDate,
                                                                           maxAmount,
                                                                           adult,
                                                                           child,
                                                                           infant,
                                                                           other,
                                                                           flightId,
                                                                           fareId,
                                                                           nonStopOnly,
                                                                           includeDeparted,
                                                                           includeCancelled,
                                                                           includeWaitlisted,
                                                                           includeSoldOut,
                                                                           includeFares,
                                                                           refundable,
                                                                           groupFares,
                                                                           iTFaresOnly,
                                                                           staffFares,
                                                                           showClose,
                                                                           unknownTransit,
                                                                           mapWithFares,
                                                                           noVat,
                                                                           transactionReference);

            if (returnFromDate != DateTime.MinValue && returnToDate != DateTime.MinValue)
            {
                //Find return availability
                availabilities.FlightAvailabilityReturn = GetAvailableFlight(otherPassengerType,
                                                                             boardingClass,
                                                                             bookingClass,
                                                                             dayTimeIndicator,
                                                                             agencyCode,
                                                                             currencyCode,
                                                                             transitPoint,
                                                                             promotionCode,
                                                                             fareType,
                                                                             languageCode,
                                                                             ipAddress,
                                                                             destinationRcd,
                                                                             originRcd,
                                                                             odDestinationRcd,
                                                                             odOriginRcd,
                                                                             returnFromDate,
                                                                             returnToDate,
                                                                             bookingDate,
                                                                             maxAmount,
                                                                             adult,
                                                                             child,
                                                                             infant,
                                                                             other,
                                                                             flightId,
                                                                             fareId,
                                                                             nonStopOnly,
                                                                             includeDeparted,
                                                                             includeCancelled,
                                                                             includeWaitlisted,
                                                                             includeSoldOut,
                                                                             includeFares,
                                                                             refundable,
                                                                             groupFares,
                                                                             iTFaresOnly,
                                                                             staffFares,
                                                                             showClose,
                                                                             unknownTransit,
                                                                             mapWithFares,
                                                                             noVat,
                                                                             transactionReference);
            }
            
            //Apply fare logic.
            if (applyFareLogic == true)
            {
 
            }


            return availabilities;
        }

        private IEnumerable<Availability> GetAvailableFlight(string otherPassengerType,
                                                            string boardingClass,
                                                            string bookingClass,
                                                            string dayTimeIndicator,
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
                                                            DateTime bookingDate,
                                                            decimal maxAmount,
                                                            byte adult,
                                                            byte child,
                                                            byte infant,
                                                            byte other,
                                                            Guid flightId,
                                                            Guid fareId,
                                                            Int16 nonStopOnly,
                                                            Int16 includeDeparted,
                                                            Int16 includeCancelled,
                                                            Int16 includeWaitlisted,
                                                            Int16 includeSoldOut,
                                                            Int16 includeFares,
                                                            Int16 refundable,
                                                            Int16 groupFares,
                                                            Int16 iTFaresOnly,
                                                            Int16 staffFares,
                                                            bool showClose,
                                                            Int16 unknownTransit,
                                                            bool mapWithFares,
                                                            bool noVat,
                                                            string transactionReference)
        {

            try
            {
                IAvailabilityBase availability = null;
                //Call base availability.

                if (_availabilityRepository != null && _flightRepository != null)
                {
                    //Find available route
                    IEnumerable<AvailabilityRoute> availabilityRoute = _flightRepository.GetFlightAvailabilityRoute(originRcd, destinationRcd);
                    //Get Base Availability
                    availability = new AvailabilityBase(_availabilityRepository,
                                                        availabilityRoute,
                                                        otherPassengerType,
                                                        boardingClass,
                                                        bookingClass,
                                                        dayTimeIndicator,
                                                        agencyCode,
                                                        currencyCode,
                                                        promotionCode,
                                                        fareType,
                                                        languageCode,
                                                        ipAddress,
                                                        originRcd,
                                                        destinationRcd,
                                                        odOriginRcd,
                                                        odDestinationRcd,
                                                        fromDate,
                                                        toDate,
                                                        bookingDate,
                                                        maxAmount,
                                                        mapWithFares,
                                                        showClose,
                                                        adult,
                                                        child,
                                                        infant,
                                                        other,
                                                        nonStopOnly,
                                                        includeDeparted,
                                                        includeCancelled,
                                                        includeWaitlisted,
                                                        includeSoldOut,
                                                        includeFares,
                                                        refundable,
                                                        groupFares,
                                                        iTFaresOnly,
                                                        staffFares,
                                                        unknownTransit,
                                                        noVat,
                                                        flightId,
                                                        fareId);

                    if (nonStopOnly == 0)
                    {
                        //Get connection flight.
                        availability = new AvailabilityConnectionFlight(availability,
                                                                        _availabilityRepository,
                                                                        _flightRepository,
                                                                        availabilityRoute,
                                                                        otherPassengerType,
                                                                        boardingClass,
                                                                        bookingClass,
                                                                        dayTimeIndicator,
                                                                        agencyCode,
                                                                        currencyCode,
                                                                        transitPoint,
                                                                        promotionCode,
                                                                        fareType,
                                                                        languageCode,
                                                                        ipAddress,
                                                                        originRcd,
                                                                        destinationRcd,
                                                                        odOriginRcd,
                                                                        odDestinationRcd,
                                                                        fromDate,
                                                                        toDate,
                                                                        bookingDate,
                                                                        maxAmount,
                                                                        mapWithFares,
                                                                        showClose,
                                                                        adult,
                                                                        child,
                                                                        infant,
                                                                        other,
                                                                        nonStopOnly,
                                                                        includeDeparted,
                                                                        includeCancelled,
                                                                        includeWaitlisted,
                                                                        includeFares,
                                                                        includeFares,
                                                                        refundable,
                                                                        groupFares,
                                                                        iTFaresOnly,
                                                                        staffFares,
                                                                        unknownTransit,
                                                                        noVat,
                                                                        flightId,
                                                                        fareId);
                    }
                }
                else
                {
                    throw new NullReferenceException("One if the repository is null");
                }

                //Decorate Availability
                if (_availabilityType == AvailabilityTypes.OWN)
                {

                    //Filter search fare type.
                    switch (fareType)
                    {
                        case AvailabilityFareTypes.LOWEST:
                            availability = new AvailabilityLowestFare(availability);
                            break;
                        case AvailabilityFareTypes.LOWESTGROUP:
                            availability = new AvailabilityLowestGroup(availability);
                            break;
                        case AvailabilityFareTypes.LOWESTCLASS:
                            availability = new AvailabilityLowestClass(availability);
                            break;
                    }

                    //Fill tax information to availability.
                    availability = new AvailabilityTax(availability,
                                                        _flightRepository,
                                                        originRcd,
                                                        destinationRcd,
                                                        otherPassengerType,
                                                        currencyCode,
                                                        agencyCode,
                                                        fromDate,
                                                        toDate,
                                                        bookingDate,
                                                        adult,
                                                        child,
                                                        infant,
                                                        noVat);
                }
                else
                {
                    throw new NotImplementedException("Selected availability type is not yet implemented!");
                }

                return availability.GetAvailability();
            }
            catch
            {
                throw;
            }
        }
    }
}
