using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using Avantik.Web.Service;
using Avantik.Web.Service.Message;

namespace Avantik.Web.Service.Extension
{
    public static class AvailabilityExtension
    {
        public static IEnumerable<AvailabilityView> MappingAvailabilityView(this IEnumerable<Entity.Flight.Availability> availability)
        {
            try
            {
                if (availability != null)
                {
                    IList<AvailabilityView> objView = new List<AvailabilityView>();
                    AvailabilityView av = null;
                    foreach (Entity.Flight.Availability a in availability)
                    {
                        av = new AvailabilityView();

                        av.flight_id = a.flight_id;
                        av.transit_fare_id = a.transit_fare_id;
                        av.transit_flight_id = a.transit_flight_id;
                        av.fare_id = a.fare_id;

                        av.airline_rcd = a.airline_rcd;
                        av.flight_number = a.flight_number;
                        av.flight_status_rcd = a.flight_status_rcd;
                        av.aircraft_type_rcd = a.aircraft_type_rcd;
                        av.origin_rcd = a.origin_rcd;
                        av.destination_rcd = a.destination_rcd;
                        av.flight_comment = a.flight_comment;
                        av.operating_airline_rcd = a.operating_airline_rcd;
                        av.operating_flight_number = a.operating_flight_number;
                        av.operating_airline_name = a.operating_airline_name;
                        av.flight_information_1 = a.flight_information_1;
                        av.flight_information_2 = a.flight_information_2;
                        av.flight_information_3 = a.flight_information_3;
                        av.endorsement_text = a.endorsement_text;
                        av.restriction_text = a.restriction_text;
                        av.origin_name = a.origin_name;
                        av.destination_name = a.destination_name;
                        av.flight_check_in_status_rcd = a.flight_check_in_status_rcd;
                        av.fare_code = a.fare_code;
                        av.fare_type_rcd = a.fare_type_rcd;
                        av.booking_class_rcd = a.booking_class_rcd;
                        av.boarding_class_rcd = a.boarding_class_rcd;
                        av.nesting_string = a.nesting_string;
                        av.origin_country_rcd = a.origin_country_rcd;
                        av.destination_country_rcd = a.destination_country_rcd;
                        av.transit_country_rcd = a.transit_country_rcd;
                        av.transit_points = a.transit_points;
                        av.transit_points_name = a.transit_points_name;
                        av.transit_airport_rcd = a.transit_airport_rcd;
                        av.transit_name = a.transit_name;
                        av.transit_airline_rcd = a.transit_airline_rcd;
                        av.transit_flight_number = a.transit_flight_number;
                        av.transit_operating_airline_rcd = a.transit_operating_airline_rcd;
                        av.transit_operating_flight_number = a.transit_operating_flight_number;
                        av.transit_flight_status_rcd = a.transit_flight_status_rcd;
                        av.transit_aircraft_type_rcd = a.transit_aircraft_type_rcd;
                        av.transit_flight_comment = a.transit_flight_comment;
                        av.transit_flight_check_in_status_rcd = a.transit_flight_check_in_status_rcd;
                        av.transit_fare_code = a.transit_fare_code;
                        av.transit_fare_type_rcd = a.transit_fare_type_rcd;
                        av.transit_booking_class_rcd = a.transit_booking_class_rcd;
                        av.transit_boarding_class_rcd = a.transit_boarding_class_rcd;
                        av.currency_rcd = a.currency_rcd;

                        av.promotion_code = a.promotion_code;
                        if (!string.IsNullOrEmpty(av.promotion_code))
                            av.promotion_text = a.promotion_text;

                        av.transit_transit_points = a.transit_transit_points;
                        av.transit_transit_points_name = a.transit_transit_points_name;
                        av.transit_nesting_string = a.transit_nesting_string;
                       // av.promotion_rcd = a.promotion_rcd;
                        av.transaction_reference = a.transaction_reference;

                        av.total_departure_date = a.total_departure_date;
                        av.departure_date = a.departure_date;
                        av.arrival_date = a.arrival_date;
                        av.transit_departure_date = a.transit_departure_date;
                        av.transit_arrival_date = a.transit_arrival_date;
                        av.utc_departure_date_time = a.utc_departure_date_time;
                        av.utc_arrival_date_time = a.utc_arrival_date_time;

                        av.planned_departure_time = a.planned_departure_time;
                        av.estimated_departure_time = a.estimated_departure_time;
                        av.planned_arrival_time = a.planned_arrival_time;
                        av.flight_duration = a.flight_duration;
                        av.number_of_stops = a.number_of_stops;
                        av.departed = a.departed;
                        av.transit_planned_departure_time = a.transit_planned_departure_time;
                        av.transit_estimated_departure_time = a.transit_estimated_departure_time;
                        av.transit_planned_arrival_time = a.transit_planned_arrival_time;
                        av.time_diff = a.time_diff;
                        av.close_web_sales = a.close_web_sales;
                        av.total_flight_duration = a.total_flight_duration;

                        av.it_fare_flag = a.it_fare_flag;
                        av.class_open_flag = a.class_open_flag;
                        av.waitlist_open_flag = a.waitlist_open_flag;
                        av.ignore_logic_flag = a.ignore_logic_flag;
                        av.transit_it_fare_flag = a.transit_it_fare_flag;
                        av.transit_class_open_flag = a.transit_class_open_flag;
                        av.transit_waitlist_open_flag = a.transit_waitlist_open_flag;
                        av.transit_ignore_logic_flag = a.transit_ignore_logic_flag;
                        av.total_class_open_flag = a.total_class_open_flag;
                        av.total_waitlist_open_flag = a.total_waitlist_open_flag;
                        av.corporate_fare_flag = a.corporate_fare_flag;
                        av.refundable_flag = a.refundable_flag;
                        av.full_flight_flag = a.full_flight_flag;
                        av.no_waitlist_flight_flag = a.no_waitlist_flight_flag;
                        av.departed_flight_flag = a.departed_flight_flag;
                        av.public_fare_flag = a.public_fare_flag;
                        av.open_return_fare_flag = a.open_return_fare_flag;
                        av.staff_fare_flag = a.staff_fare_flag;
                        av.subload_fare_flag = a.subload_fare_flag;
                        av.e_ticket_flag = a.e_ticket_flag;
                        av.avl_show_net_total_flag = a.avl_show_net_total_flag;

                        av.filter_logic_flag = a.filter_logic_flag;
                        av.class_capacity = a.class_capacity;
                        av.pax_book_count = a.pax_book_count;
                        av.pax_waitlist_count = a.pax_waitlist_count;
                        av.waitlist_capacity = a.waitlist_capacity;
                        av.close_hours = a.close_hours;
                        av.reopen_hours = a.reopen_hours;
                        av.nested_book_available = a.nested_book_available;
                        av.physical_capacity = a.physical_capacity;
                        av.bookable_capacity = a.bookable_capacity;
                        av.transit_flight_duration = a.transit_flight_duration;
                        av.transit_number_of_stops = a.transit_number_of_stops;
                        av.transit_departed = a.transit_departed;
                        av.transit_class_capacity = a.transit_class_capacity;
                        av.transit_pax_book_count = a.transit_pax_book_count;
                        av.transit_pax_waitlist_count = a.transit_pax_waitlist_count;
                        av.transit_waitlist_capacity = a.transit_waitlist_capacity;
                        av.transit_close_hours = a.transit_close_hours;
                        av.transit_reopen_hours = a.transit_reopen_hours;
                        av.transit_nested_book_available = a.transit_nested_book_available;
                        av.transit_physical_capacity = a.transit_physical_capacity;
                        av.transit_bookable_capacity = a.transit_bookable_capacity;
                        av.total_planned_departure_time = a.total_planned_departure_time;
                        av.total_planned_arrival_time = a.total_planned_arrival_time;
                        av.total_number_of_stops = a.total_number_of_stops;
                        av.total_class_capacity = a.total_class_capacity;
                        av.total_pax_book_count = a.total_pax_book_count;
                        av.total_pax_waitlist_count = a.total_pax_waitlist_count;
                        av.total_waitlist_capacity = a.total_waitlist_capacity;
                        av.total_nested_book_available = a.total_nested_book_available;
                        av.total_physical_capacity = a.total_physical_capacity;
                        av.total_bookable_capacity = a.total_bookable_capacity;
                        av.fare_column = a.fare_column;
                        av.timelimit_hours_before_flight = a.timelimit_hours_before_flight;
                        av.timelimit_hours_after_booking = a.timelimit_hours_after_booking;

                        av.redemption_points = a.redemption_points;
                        av.adult_fare = a.adult_fare;
                        av.child_fare = a.child_fare;
                        av.infant_fare = a.infant_fare;
                        av.other_fare = a.other_fare;
                        av.transit_redemption_points = a.transit_redemption_points;
                        av.transit_adult_fare = a.transit_adult_fare;
                        av.transit_child_fare = a.transit_child_fare;
                        av.transit_infant_fare = a.transit_infant_fare;
                        av.transit_other_fare = a.transit_other_fare;
                        av.total_redemption_points = a.total_redemption_points;
                        av.total_adult_fare = a.total_adult_fare;
                        av.total_child_fare = a.total_child_fare;
                        av.total_infant_fare = a.total_infant_fare;
                        av.total_other_fare = a.total_other_fare;
                        av.adult_fare_excl = a.adult_fare_excl;
                        av.child_fare_excl = a.child_fare_excl;
                        av.infant_fare_excl = a.infant_fare_excl;
                        av.other_fare_excl = a.other_fare_excl;

                        objView.Add(av);

                    }

                    return objView;
                }
                else
                {
                    return null;
                }
            }
            catch
            {
                throw;
            }
        }
    }
}
