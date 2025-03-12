using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using Avantik.Web.Service.Entity.Flight;
using Avantik.Web.Service.Repository.Contract.Flight;
using Avantik.Web.Service.Infrastructure;

namespace Avantik.Web.Service.Model
{
    public class AvailabilityBase : IAvailabilityBase
    {
        protected IAvailabilityRepository _availabilityRepository;
        protected IEnumerable<AvailabilityRoute> _availabilityRoute;

        //Search Flight Argement
        string _otherPassengerType;
        string _boardingClass;
        string _bookingClass;
        string _dayTimeIndicator;
        string _agencyCode;
        string _currencyCode;
        string _promotionCode;
        AvailabilityFareTypes _fareType;
        string _languageCode;
        string _ipAddress;
        string _originRcd;
        string _destinationRcd;
        string _odOriginRcd;
        string _odDestinationRcd;

        DateTime _fromDate;
        DateTime _toDate;
        DateTime _bookingDate;

        bool _mapWithFares;
        bool _applyFareLogic;
        bool _showClose;
        bool _noVat;

        decimal _maxAmount;

        byte _adult;
        byte _child;
        byte _infant;
        byte _other;

        Int16 _nonStopOnly;
        Int16 _includeDeparted;
        Int16 _includeCancelled;
        Int16 _includeWaitlisted;
        Int16 _includeSoldOut;
        Int16 _includeFares;
        Int16 _refundable;
        Int16 _groupFares;
        Int16 _iTFaresOnly;
        Int16 _staffFares;
        Int16 _unknownTransit;
        

        Guid _flightId;
        Guid _fareId;

        public AvailabilityBase(IAvailabilityRepository availabilityRepository,
                                IEnumerable<AvailabilityRoute> availabilityRoute,
                                string otherPassengerType,
                                string boardingClass,
                                string bookingClass,
                                string dayTimeIndicator,
                                string agencyCode,
                                string currencyCode,
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
                                bool mapWithFares,
                                bool showClose,
                                byte adult,
                                byte child,
                                byte infant,
                                byte other,
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
                                Int16 unknownTransit,
                                bool noVat,
                                Guid flightId,
                                Guid fareId)
        {
            _availabilityRepository = availabilityRepository;
            _availabilityRoute = availabilityRoute;

            //Assign Parameter.
            _otherPassengerType = otherPassengerType;
            _boardingClass = boardingClass;
            _bookingClass = bookingClass;
            _dayTimeIndicator = dayTimeIndicator;
            _agencyCode = agencyCode;
            _currencyCode = currencyCode;
            _promotionCode = promotionCode;
            _fareType = fareType;
            _languageCode = languageCode;
            _ipAddress = ipAddress;
            _originRcd = originRcd;
            _destinationRcd = destinationRcd;
            _odOriginRcd = odOriginRcd;
            _odDestinationRcd = odDestinationRcd;

            _fromDate = fromDate;
            _toDate = toDate;
            _bookingDate = bookingDate;

            _maxAmount = maxAmount;

            _adult = adult;
            _child = child;
            _infant = infant;
            _other = other;

            _flightId = flightId;
            _fareId = fareId;

            _nonStopOnly = nonStopOnly;
            _includeDeparted = includeDeparted;
            _includeCancelled = includeCancelled;
            _includeWaitlisted = includeWaitlisted;
            _includeSoldOut = includeSoldOut;
            _includeFares = includeFares;
            _refundable = refundable;
            _groupFares = groupFares;
            _iTFaresOnly = iTFaresOnly;
            _staffFares = staffFares;
            _unknownTransit = unknownTransit;

            _showClose = showClose;
            _mapWithFares = mapWithFares;
            
            _noVat = noVat;
        }

        public IList<Availability> GetAvailability()
        {
            IList<Availability> availabilityResult = null;
            try
            {
                if (_availabilityRepository != null)
                {
                    if (_availabilityRoute != null)
                    {
                        //Get Base availability.
                        IEnumerable<Availability> availability;
                        availabilityResult = new List<Availability>();

                        foreach (AvailabilityRoute r in _availabilityRoute)
                        {
                            if (r.transit_airport_rcd == string.Empty &&
                                r.direct_flag != 0)
                            {
                                //Find the main flight legs.
                                availability = _availabilityRepository.FindAvailability(r.avl_origin_rcd,
                                                                                        r.avl_destination_rcd,
                                                                                        _bookingClass,
                                                                                        _boardingClass,
                                                                                        _otherPassengerType,
                                                                                        _agencyCode,
                                                                                        _currencyCode,
                                                                                        _dayTimeIndicator,
                                                                                        _odOriginRcd,
                                                                                        _odDestinationRcd,
                                                                                        _promotionCode,
                                                                                        _fareType,
                                                                                        _languageCode,
                                                                                        _ipAddress,
                                                                                        _fromDate,
                                                                                        _toDate,
                                                                                        _bookingDate,
                                                                                        _flightId,
                                                                                        _fareId,
                                                                                        _adult,
                                                                                        _child,
                                                                                        _infant,
                                                                                        _other,
                                                                                        _nonStopOnly,
                                                                                        _groupFares,
                                                                                        _refundable,
                                                                                        _includeCancelled,
                                                                                        _includeDeparted,
                                                                                        _includeWaitlisted,
                                                                                        _includeSoldOut,
                                                                                        _iTFaresOnly,
                                                                                        _staffFares,
                                                                                        _includeFares,
                                                                                        _maxAmount,
                                                                                        _mapWithFares,
                                                                                        string.Empty);

                                if (availability != null)
                                {
                                    //Fill main flight leg to availability result object.
                                    foreach (Availability a in availability)
                                    {

                                        a.SetTotalClassOpenFlag();
                                        a.SetTotalWailistOpenFlag();
                                        a.SetTotalNestBookAvai();
                                        a.SetTotalPhysicalCapacity();
                                        a.SetTotalBookableCapacity();
                                        a.total_redemption_points = a.redemption_points + a.transit_redemption_points;

                                        //Set full flight flag.
                                        a.SetFullFlightFlag(_adult, _child, _other);

                                        //Set Notvat
                                        a.SetVAT(_noVat);

                                        //Set Total fare
                                        a.SetTotalFare();

                                        if (a.FilterOutAvailability(_adult, _child, _other, _showClose) == false)
                                        {
                                            availabilityResult.Add(a);
                                        }
                                        
                                    }
                                }
                            }
                        }

                    }
                }
            }
            catch (Exception ex)
            {
                throw;
            }

            return availabilityResult;
        }
    }
}
