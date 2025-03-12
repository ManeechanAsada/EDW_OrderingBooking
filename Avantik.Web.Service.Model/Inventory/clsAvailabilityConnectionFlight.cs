using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using Avantik.Web.Service.Entity.Flight;
using Avantik.Web.Service.Repository.Contract.Flight;
using Avantik.Web.Service.Infrastructure;

namespace Avantik.Web.Service.Model
{
    public class AvailabilityConnectionFlight : AvailabilityDecorator
    {
        protected IAvailabilityRepository _availabilityRepository;
        protected IFlightRepository _flightRepository;

        protected IEnumerable<AvailabilityRoute> _availabilityRoute;

        //Search Flight Argement
        string _otherPassengerType;
        string _boardingClass;
        string _bookingClass;
        string _dayTimeIndicator;
        string _agencyCode;
        string _currencyCode;
        string _transitPoint;
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

        public AvailabilityConnectionFlight(IAvailabilityBase availability, 
                                            IAvailabilityRepository availabilityRepository,
                                            IFlightRepository flightRepository,
                                            IEnumerable<AvailabilityRoute> availabilityRoute,
                                            string otherPassengerType,
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
            : base(availability)
        {
            
            _availabilityRepository = availabilityRepository;
            _flightRepository = flightRepository;

            _availabilityRoute = availabilityRoute;

            //Assign Parameter.
            _otherPassengerType = otherPassengerType;
            _boardingClass = boardingClass;
            _bookingClass = bookingClass;
            _dayTimeIndicator = dayTimeIndicator;
            _agencyCode = agencyCode;
            _currencyCode = currencyCode;
            _transitPoint = transitPoint;
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

            _mapWithFares = mapWithFares;
            _showClose = showClose;

            _noVat = noVat;
        }

        public override IList<Availability> GetAvailability()
        {
            try
            {
                IList<Availability> baseAvailability = base._Availability.GetAvailability();
                IEnumerable<Availability> firstAvailability = null;
                IEnumerable<Availability> secondAvailability = null;
                IEnumerable<AvailabilityParing> availabilityParing;

                DateTime dtMin = DateTime.MinValue;
                DateTime dtMax = DateTime.MinValue;

                if (baseAvailability == null)
                {
                    baseAvailability = new List<Availability>();
                }

                //Find connection flight availability.
                if (_availabilityRoute != null)
                {

                    string strFirstFrom;
                    string strFirstTo;

                    string strSecondFrom;
                    string strSecondTo;

                    foreach (AvailabilityRoute r in _availabilityRoute)
                    {
                        //Reset airprot information.
                        strFirstFrom = "";
                        strFirstTo = "";

                        strSecondFrom = "";
                        strSecondTo = "";

                        //Find airport information
                        if (string.IsNullOrEmpty(_transitPoint))
                        {
                            if (string.IsNullOrEmpty(r.transit_airport_rcd) == false &&
                                r.direct_flag != 0)
                            {
                                strFirstFrom = r.avl_origin_rcd;
                                strFirstTo = r.transit_airport_rcd;

                                strSecondFrom = r.transit_airport_rcd;
                                strSecondTo = r.avl_destination_rcd;
                            }
                        }
                        else
                        {
                            if (r.transit_airport_rcd == _transitPoint &&
                                r.direct_flag != 0)
                            {
                                strFirstFrom = r.avl_origin_rcd;
                                strFirstTo = r.transit_airport_rcd;

                                strSecondFrom = r.transit_airport_rcd;
                                strSecondTo = r.avl_destination_rcd;
                            }
                        }

                        //Get transit availability for each alternative.
                        if (r.dynamic_connections_flag == 1)
                        {
                            availabilityParing = null;
                        }
                        else if (r.connex > 0)
                        {
                            availabilityParing = _flightRepository.GetFlightAvailabilityParing(r.origin_rcd,
                                                                                                r.destination_rcd,
                                                                                                r.transit_airport_rcd,
                                                                                                _fromDate,
                                                                                                _toDate);
                        }
                        else
                        {
                            availabilityParing = null;
                        }

                        //Get availability from repository.
                        if (string.IsNullOrEmpty(strFirstFrom) == false && string.IsNullOrEmpty(strFirstTo) == false)
                        {
                            //Find availability for the first leg.
                            firstAvailability = _availabilityRepository.FindAvailability(strFirstFrom,
                                                                                         strFirstTo,
                                                                                         _bookingClass,
                                                                                         _boardingClass,
                                                                                         _otherPassengerType,
                                                                                         _agencyCode,
                                                                                         _currencyCode,
                                                                                         _dayTimeIndicator,
                                                                                         _originRcd,
                                                                                         _destinationRcd,
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

                            if (string.IsNullOrEmpty(strSecondFrom) == false && string.IsNullOrEmpty(strSecondTo) == false)
                            {
                                //Find availability for the second leg.
                                secondAvailability = _availabilityRepository.FindAvailability(strSecondFrom,
                                                                                              strSecondTo,
                                                                                              _bookingClass,
                                                                                              _boardingClass,
                                                                                              _otherPassengerType,
                                                                                              _agencyCode,
                                                                                              _currencyCode,
                                                                                              _dayTimeIndicator,
                                                                                              _originRcd,
                                                                                              _destinationRcd,
                                                                                              _promotionCode,
                                                                                              _fareType,
                                                                                              _languageCode,
                                                                                              _ipAddress,
                                                                                              _fromDate,
                                                                                              _toDate.AddDays(1),
                                                                                              _bookingDate,
                                                                                              _flightId,
                                                                                              _fareId,
                                                                                              _adult,
                                                                                              _child,
                                                                                              _infant,
                                                                                              _other,
                                                                                              0,                  //Set "0" because the base availability will alway search for short leg.
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

                            }

                            //Combine flight.
                            if (firstAvailability != null && secondAvailability != null)
                            {
                                foreach (Availability fa in firstAvailability)
                                {
                                    if (availabilityParing != null)
                                    {
                                        //Manual paring
                                        dtMin = fa.arrival_date.AddMinutes(0);
                                        dtMax = fa.arrival_date.AddMinutes(600);

                                        foreach (AvailabilityParing ap in availabilityParing)
                                        {
                                            if (ap.leg_1_airline_rcd == fa.airline_rcd && ap.leg_1_flight_number == fa.flight_number)
                                            {
                                                foreach (Availability fs in secondAvailability)
                                                {
                                                    if (fs.airline_rcd == ap.leg_2_airline_rcd &&
                                                        fs.flight_number == ap.leg_2_flight_number &&
                                                        fs.booking_class_rcd == fa.booking_class_rcd &&
                                                        fs.fare_id == fa.fare_id &&
                                                        fs.utc_departure_date_time >= dtMin &&
                                                        fs.utc_departure_date_time <= dtMax &&
                                                        fs.flight_id != fa.flight_id)
                                                    {
                                                        //Start Flight combine.
                                                        CombineTransitFlight(baseAvailability, fa, fs);
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    else
                                    {
                                        dtMin = fa.arrival_date.AddMinutes(r.min_transit_minutes);
                                        dtMax = fa.arrival_date.AddMinutes(r.max_transit_minutes);
                                        foreach (Availability fs in secondAvailability)
                                        {
                                            if (fs.booking_class_rcd == fa.booking_class_rcd &&
                                                fs.fare_id == fa.fare_id &&
                                                fs.utc_departure_date_time >= dtMin &&
                                                fs.utc_departure_date_time <= dtMax &&
                                                fs.flight_id != fa.flight_id)
                                            {
                                                //Start Flight combine.
                                                CombineTransitFlight(baseAvailability, fa, fs);
                                            }
                                        }
                                    }
                                }
                            }
                        }

                    }
                }
                return baseAvailability;
            }
            catch (Exception ex)
            {
                throw;
            }
        }
        
        private void CombineTransitFlight(IList<Availability> availability,
                                          Availability first,
                                          Availability second)
        {
            Availability a;

            if (availability != null)
            {
                a = new Availability();

                a.avl_show_net_total_flag = first.avl_show_net_total_flag;
                a.flight_id = first.flight_id;
                a.airline_rcd = first.airline_rcd;
                a.flight_number = first.flight_number;
                a.operating_airline_rcd = first.operating_airline_rcd;
                a.operating_flight_number = first.operating_flight_number;
                a.operating_airline_name = first.operating_airline_name;
                a.flight_status_rcd = first.flight_status_rcd;
                a.aircraft_type_rcd = first.aircraft_type_rcd;
                a.flight_information_1 = first.flight_information_1;
                a.flight_information_2 = first.flight_information_2;
                a.flight_information_3 = first.flight_information_3;
                a.endorsement_text = first.endorsement_text;
                a.restriction_text = first.restriction_text;
                a.flight_check_in_status_rcd = first.flight_check_in_status_rcd;
                a.departure_date = first.departure_date;
                a.planned_departure_time = first.planned_departure_time;
                a.estimated_departure_time = first.estimated_departure_time;
                a.planned_arrival_time = first.planned_arrival_time;
                a.arrival_date = first.arrival_date;
                a.flight_duration = first.flight_duration;
                a.number_of_stops = first.number_of_stops;
                a.departed = first.departed;
                a.fare_code = first.fare_code;
                a.fare_type_rcd = first.fare_type_rcd;
                a.fare_id = first.fare_id;
                a.it_fare_flag = first.it_fare_flag;
                a.corporate_fare_flag = first.corporate_fare_flag;
                a.refundable_flag = first.refundable_flag;
               // a.promotion_code = first.promotion_rcd;
                a.promotion_text = first.promotion_text;
                a.booking_class_rcd = first.booking_class_rcd;
                a.boarding_class_rcd = first.boarding_class_rcd;
                a.class_open_flag = first.class_open_flag;
                a.class_capacity = first.class_capacity;
                a.pax_book_count = first.pax_book_count;
                a.pax_waitlist_count = first.pax_waitlist_count;
                a.waitlist_capacity = first.waitlist_capacity;
                a.waitlist_open_flag = first.waitlist_open_flag;
                a.nesting_string = first.nesting_string;
                a.close_hours = first.close_hours;
                a.reopen_hours = first.reopen_hours;
                a.ignore_logic_flag = first.ignore_logic_flag;
                a.nested_book_available = first.nested_book_available;
                a.physical_capacity = first.physical_capacity;
                a.bookable_capacity = first.bookable_capacity;
                a.redemption_points = first.redemption_points;
                a.origin_country_rcd = first.origin_country_rcd;
                a.destination_country_rcd = first.destination_country_rcd;
                a.transit_country_rcd = first.transit_country_rcd;
                a.transit_points = first.transit_points;
                a.transit_points_name = first.transit_points_name;
                a.currency_rcd = first.currency_rcd;

                if (string.IsNullOrEmpty(first.flight_comment) == false)
                {
                    if (first.flight_comment.Length > 60)
                    {
                        a.flight_comment = first.flight_comment.Substring(0, 60);
                    }
                    else
                    {
                        a.flight_comment = first.flight_comment;
                    }
                }

                if (_noVat == true)
                {
                    a.adult_fare = Math.Round(first.adult_fare_excl, 2);
                    a.child_fare = Math.Round(first.child_fare_excl, 2);
                    a.infant_fare = Math.Round(first.infant_fare_excl, 2);
                    a.other_fare = Math.Round(first.other_fare_excl, 2);
                }
                else
                {
                    a.adult_fare = Math.Round(first.adult_fare, 2);
                    a.child_fare = Math.Round(first.child_fare, 2);
                    a.infant_fare = Math.Round(first.infant_fare, 2);
                    a.other_fare = Math.Round(first.other_fare, 2);
                }

                //-------------------------------------------------------------------------
                //combine flight - leg 2
                //-------------------------------------------------------------------------
                a.transit_flight_id = second.flight_id;
                a.transit_airline_rcd = second.airline_rcd;
                a.transit_flight_number = second.flight_number;
                a.transit_operating_airline_rcd = second.operating_airline_rcd;
                a.transit_operating_flight_number = second.operating_flight_number;
                
                a.transit_flight_status_rcd = second.flight_status_rcd;
                a.transit_aircraft_type_rcd = second.aircraft_type_rcd;
                a.transit_flight_check_in_status_rcd = second.flight_check_in_status_rcd;
                a.transit_departure_date = second.departure_date;
                a.transit_planned_departure_time = second.planned_departure_time;
                a.transit_estimated_departure_time = second.estimated_departure_time;
                a.transit_planned_arrival_time = second.planned_arrival_time;
                a.transit_arrival_date = second.arrival_date;
                a.transit_flight_duration = second.flight_duration;
                a.transit_number_of_stops = second.number_of_stops;
                a.transit_departed = second.departed;
                a.transit_fare_code = second.fare_code;
                a.transit_fare_type_rcd = second.fare_type_rcd;
                a.transit_fare_id = second.fare_id;
                a.transit_it_fare_flag = second.it_fare_flag;
                a.transit_booking_class_rcd = second.booking_class_rcd;
                a.transit_boarding_class_rcd = second.boarding_class_rcd;
                a.transit_class_open_flag = second.class_open_flag;
                a.transit_class_capacity = second.class_capacity;
                a.transit_pax_book_count = second.pax_book_count;
                a.transit_pax_waitlist_count = second.pax_waitlist_count;
                a.transit_waitlist_capacity = second.waitlist_capacity;
                a.transit_waitlist_open_flag = second.waitlist_open_flag;
                a.transit_nesting_string = second.nesting_string;
                a.transit_close_hours = second.close_hours;
                a.transit_reopen_hours = second.reopen_hours;
                a.transit_ignore_logic_flag = second.ignore_logic_flag;
                a.transit_nested_book_available = second.nested_book_available;
                a.transit_physical_capacity = second.physical_capacity;
                a.transit_bookable_capacity = second.bookable_capacity;
                a.transit_redemption_points = Math.Round(second.redemption_points, 2);
                a.transit_transit_points = second.transit_points;
                a.transit_transit_points_name = second.transit_points_name;
                
                if (string.IsNullOrEmpty(second.flight_comment) == false)
                {
                    if (second.flight_comment.Length > 60)
                    {
                        a.transit_flight_comment = second.flight_comment.Substring(0, 60);
                    }
                    else
                    {
                        a.transit_flight_comment = second.flight_comment;
                    }
                }

                if (_noVat == true)
                {
                    a.transit_adult_fare = Math.Round(second.adult_fare_excl, 2);
                    a.transit_child_fare = Math.Round(second.child_fare_excl, 2);
                    a.transit_infant_fare = Math.Round(second.infant_fare_excl, 2);
                    a.transit_other_fare = Math.Round(second.other_fare_excl, 2);
                }
                else
                {
                    a.transit_adult_fare = Math.Round(second.adult_fare, 2);
                    a.transit_child_fare = Math.Round(second.child_fare, 2);
                    a.transit_infant_fare = Math.Round(second.infant_fare, 2);
                    a.transit_other_fare = Math.Round(second.other_fare, 2);
                }

                a.origin_rcd = first.origin_rcd;
                a.origin_name = first.origin_name;
                a.transit_airport_rcd = first.destination_rcd;
                a.transit_name = first.destination_name;
                a.destination_rcd = second.destination_rcd;
                a.destination_name = second.destination_name;
                //-------------------------------------------------------------------------
                //combine flight - totals
                //-------------------------------------------------------------------------
                a.total_departure_date = a.departure_date;
                a.total_planned_departure_time = a.planned_departure_time;
                a.total_planned_arrival_time = a.transit_planned_arrival_time;
                a.filter_logic_flag = -1;
                a.e_ticket_flag = first.e_ticket_flag;

                a.time_diff = first.time_diff;
                a.close_web_sales = first.close_web_sales;
                a.staff_fare_flag = first.staff_fare_flag;
                a.subload_fare_flag = first.subload_fare_flag;
                a.open_return_fare_flag = first.open_return_fare_flag;
                a.fare_column = first.fare_column;
                a.timelimit_hours_before_flight = first.timelimit_hours_before_flight;
                a.timelimit_hours_after_booking = first.timelimit_hours_after_booking;

                a.total_flight_duration = Avantik.Web.Service.Helpers.Date.DateDiffMinute(first.utc_departure_date_time, second.utc_arrival_date_time);

                a.total_number_of_stops = a.number_of_stops + a.transit_number_of_stops + 1;
                
                a.SetTotalTransitClassCapacity();
                a.SetTotalTransitWaitlistCapacity();
                a.SetTotalClassOpenFlag();
                a.SetTotalWailistOpenFlag();
                a.SetTotalNestBookAvai();
                a.SetTotalPhysicalCapacity();
                a.SetTotalBookableCapacity(); 
                a.total_redemption_points = a.redemption_points + a.transit_redemption_points;
                a.SetFullFlightFlag(_adult, _child, _other);

                //Set Notvat
                a.SetVAT(_noVat);

                //Set Total fare
                a.SetTotalFare();

                if (a.FilterOutAvailability(_adult, _child, _other, _showClose) == false)
                {
                    //Fill availability object
                    availability.Add(a);
                }
            }
        }
    }
}
