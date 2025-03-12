using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using Avantik.Web.Service.Entity.Flight;
using Avantik.Web.Service.COMHelper;

namespace Avantik.Web.Service.Model.COM
{
    public static class RecordsetObjectMapping
    {
        #region Availability
        public static void FillAvailability(this IList<Availability> availability, ADODB.Recordset rs, string avalabilityLeg)
        {
            if (availability != null)
            {
                if (rs != null && rs.RecordCount > 0)
                {
                    Availability av = null;

                    rs.MoveFirst();
                    while (!rs.EOF)
                    {
                        av = new Availability();

                        av.AvalabilityLeg = avalabilityLeg;
                        av.flight_id = RecordsetHelper.ToGuid(rs, "flight_id");
                        av.transit_fare_id = RecordsetHelper.ToGuid(rs, "transit_fare_id");
                        av.transit_flight_id = RecordsetHelper.ToGuid(rs, "transit_flight_id");
                        av.fare_id = RecordsetHelper.ToGuid(rs, "fare_id");

                        av.airline_rcd = RecordsetHelper.ToString(rs, "airline_rcd");
                        av.flight_number = RecordsetHelper.ToString(rs, "flight_number");
                        av.flight_status_rcd = RecordsetHelper.ToString(rs, "flight_status_rcd");
                        av.aircraft_type_rcd = RecordsetHelper.ToString(rs, "aircraft_type_rcd");
                        av.origin_rcd = RecordsetHelper.ToString(rs, "origin_rcd");
                        av.destination_rcd = RecordsetHelper.ToString(rs, "destination_rcd");
                        av.flight_comment = RecordsetHelper.ToString(rs, "flight_comment");
                        av.operating_airline_rcd = RecordsetHelper.ToString(rs, "operating_airline_rcd");
                        av.operating_flight_number = RecordsetHelper.ToString(rs, "operating_flight_number");
                        av.operating_airline_name = RecordsetHelper.ToString(rs, "operating_airline_name");
                        av.flight_information_1 = RecordsetHelper.ToString(rs, "flight_information_1");
                        av.flight_information_2 = RecordsetHelper.ToString(rs, "flight_information_2");
                        av.flight_information_3 = RecordsetHelper.ToString(rs, "flight_information_3");
                        av.endorsement_text = RecordsetHelper.ToString(rs, "endorsement_text");
                        av.restriction_text = RecordsetHelper.ToString(rs, "restriction_text");
                        av.origin_name = RecordsetHelper.ToString(rs, "origin_name");
                        av.destination_name = RecordsetHelper.ToString(rs, "destination_name");
                        av.flight_check_in_status_rcd = RecordsetHelper.ToString(rs, "flight_check_in_status_rcd");
                        av.fare_code = RecordsetHelper.ToString(rs, "fare_code");
                        av.fare_type_rcd = RecordsetHelper.ToString(rs, "fare_type_rcd");
                        av.booking_class_rcd = RecordsetHelper.ToString(rs, "booking_class_rcd");
                        av.boarding_class_rcd = RecordsetHelper.ToString(rs, "boarding_class_rcd");
                        av.nesting_string = RecordsetHelper.ToString(rs, "nesting_string");
                        av.origin_country_rcd = RecordsetHelper.ToString(rs, "origin_country_rcd");
                        av.destination_country_rcd = RecordsetHelper.ToString(rs, "destination_country_rcd");
                        av.transit_country_rcd = RecordsetHelper.ToString(rs, "transit_country_rcd");
                        av.transit_points = RecordsetHelper.ToString(rs, "transit_points");
                        av.transit_points_name = RecordsetHelper.ToString(rs, "transit_points_name");
                        av.transit_airport_rcd = RecordsetHelper.ToString(rs, "transit_airport_rcd");
                        av.transit_name = RecordsetHelper.ToString(rs, "transit_name");
                        av.transit_airline_rcd = RecordsetHelper.ToString(rs, "transit_airline_rcd");
                        av.transit_flight_number = RecordsetHelper.ToString(rs, "transit_flight_number");
                        av.transit_operating_airline_rcd = RecordsetHelper.ToString(rs, "transit_operating_airline_rcd");
                        av.transit_operating_flight_number = RecordsetHelper.ToString(rs, "transit_operating_flight_number");
                        av.transit_flight_status_rcd = RecordsetHelper.ToString(rs, "transit_flight_status_rcd");
                        av.transit_aircraft_type_rcd = RecordsetHelper.ToString(rs, "transit_aircraft_type_rcd");
                        av.transit_flight_comment = RecordsetHelper.ToString(rs, "transit_flight_comment");
                        av.transit_flight_check_in_status_rcd = RecordsetHelper.ToString(rs, "transit_flight_check_in_status_rcd");
                        av.transit_fare_code = RecordsetHelper.ToString(rs, "transit_fare_code");
                        av.transit_fare_type_rcd = RecordsetHelper.ToString(rs, "transit_fare_type_rcd");
                        av.transit_booking_class_rcd = RecordsetHelper.ToString(rs, "transit_booking_class_rcd");
                        av.transit_boarding_class_rcd = RecordsetHelper.ToString(rs, "transit_boarding_class_rcd");
                        av.currency_rcd = RecordsetHelper.ToString(rs, "currency_rcd");
                        av.promotion_code = RecordsetHelper.ToString(rs, "promotion_code");
                        av.promotion_text = RecordsetHelper.ToString(rs, "promotion_text");
                        av.transit_transit_points = RecordsetHelper.ToString(rs, "transit_transit_points");
                        av.transit_transit_points_name = RecordsetHelper.ToString(rs, "transit_transit_points_name");
                        av.transit_nesting_string = RecordsetHelper.ToString(rs, "transit_nesting_string");
                      //  av.promotion_rcd = RecordsetHelper.ToString(rs, "promotion_rcd");
                        av.transaction_reference = RecordsetHelper.ToString(rs, "transaction_reference");

                        av.total_departure_date = RecordsetHelper.ToDateTime(rs, "total_departure_date");
                        av.departure_date = RecordsetHelper.ToDateTime(rs, "departure_date");
                        av.arrival_date = RecordsetHelper.ToDateTime(rs, "arrival_date");
                        av.transit_departure_date = RecordsetHelper.ToDateTime(rs, "transit_departure_date");
                        av.transit_arrival_date = RecordsetHelper.ToDateTime(rs, "transit_arrival_date");
                        av.utc_departure_date_time = RecordsetHelper.ToDateTime(rs, "utc_departure_date_time");
                        av.utc_arrival_date_time = RecordsetHelper.ToDateTime(rs, "utc_arrival_date_time");

                        av.planned_departure_time = RecordsetHelper.ToInt16(rs, "planned_departure_time");
                        av.estimated_departure_time = RecordsetHelper.ToInt16(rs, "estimated_departure_time");
                        av.planned_arrival_time = RecordsetHelper.ToInt16(rs, "planned_arrival_time");
                        av.flight_duration = RecordsetHelper.ToInt16(rs, "flight_duration");
                        av.number_of_stops = RecordsetHelper.ToInt16(rs, "number_of_stops");
                        av.departed = RecordsetHelper.ToInt16(rs, "departed");
                        av.transit_planned_departure_time = RecordsetHelper.ToInt16(rs, "transit_planned_departure_time");
                        av.transit_estimated_departure_time = RecordsetHelper.ToInt16(rs, "transit_estimated_departure_time");
                        av.transit_planned_arrival_time = RecordsetHelper.ToInt16(rs, "transit_planned_arrival_time");
                        av.time_diff = RecordsetHelper.ToInt32(rs, "time_diff");
                        av.close_web_sales = RecordsetHelper.ToInt32(rs, "close_web_sales");
                        av.total_flight_duration = RecordsetHelper.ToInt32(rs, "total_flight_duration");

                        av.it_fare_flag = RecordsetHelper.ToByte(rs, "it_fare_flag");
                        av.class_open_flag = RecordsetHelper.ToByte(rs, "class_open_flag");
                        av.waitlist_open_flag = RecordsetHelper.ToByte(rs, "waitlist_open_flag");
                        av.ignore_logic_flag = RecordsetHelper.ToByte(rs, "ignore_logic_flag");
                        av.transit_it_fare_flag = RecordsetHelper.ToByte(rs, "transit_it_fare_flag");
                        av.transit_class_open_flag = RecordsetHelper.ToByte(rs, "transit_class_open_flag");
                        av.transit_waitlist_open_flag = RecordsetHelper.ToByte(rs, "transit_waitlist_open_flag");
                        av.transit_ignore_logic_flag = RecordsetHelper.ToByte(rs, "transit_ignore_logic_flag");
                        av.total_class_open_flag = RecordsetHelper.ToByte(rs, "total_class_open_flag");
                        av.total_waitlist_open_flag = RecordsetHelper.ToByte(rs, "total_waitlist_open_flag");
                        av.corporate_fare_flag = RecordsetHelper.ToByte(rs, "corporate_fare_flag");
                        av.refundable_flag = RecordsetHelper.ToByte(rs, "refundable_flag");
                        av.transferable_fare_flag = RecordsetHelper.ToByte(rs, "transferable_fare_flag");
                        av.full_flight_flag = RecordsetHelper.ToByte(rs, "full_flight_flag");
                        av.no_waitlist_flight_flag = RecordsetHelper.ToByte(rs, "no_waitlist_flight_flag");
                        av.departed_flight_flag = RecordsetHelper.ToByte(rs, "departed_flight_flag");
                        av.public_fare_flag = RecordsetHelper.ToByte(rs, "public_fare_flag");
                        av.open_return_fare_flag = RecordsetHelper.ToByte(rs, "open_return_fare_flag");
                        av.staff_fare_flag = RecordsetHelper.ToByte(rs, "staff_fare_flag");
                        av.subload_fare_flag = RecordsetHelper.ToByte(rs, "subload_fare_flag");
                        av.e_ticket_flag = RecordsetHelper.ToByte(rs, "e_ticket_flag");
                        av.avl_show_net_total_flag = RecordsetHelper.ToByte(rs, "avl_show_net_total_flag");
                        av.filter_logic_flag = RecordsetHelper.ToByte(rs, "filter_logic_flag");

                        av.class_capacity = RecordsetHelper.ToInt32(rs, "class_capacity");
                        av.pax_book_count = RecordsetHelper.ToInt32(rs, "pax_book_count");
                        av.pax_waitlist_count = RecordsetHelper.ToInt32(rs, "pax_waitlist_count");
                        av.waitlist_capacity = RecordsetHelper.ToInt32(rs, "waitlist_capacity");
                        av.close_hours = RecordsetHelper.ToInt32(rs, "close_hours");
                        av.reopen_hours = RecordsetHelper.ToInt32(rs, "reopen_hours");
                        av.nested_book_available = RecordsetHelper.ToInt32(rs, "nested_book_available");
                        av.physical_capacity = RecordsetHelper.ToInt32(rs, "physical_capacity");
                        av.bookable_capacity = RecordsetHelper.ToInt32(rs, "bookable_capacity");
                        av.transit_flight_duration = RecordsetHelper.ToInt32(rs, "transit_flight_duration");
                        av.transit_number_of_stops = RecordsetHelper.ToInt32(rs, "transit_number_of_stops");
                        av.transit_departed = RecordsetHelper.ToInt32(rs, "transit_departed");
                        av.transit_class_capacity = RecordsetHelper.ToInt32(rs, "transit_class_capacity");
                        av.transit_pax_book_count = RecordsetHelper.ToInt32(rs, "transit_pax_book_count");
                        av.transit_pax_waitlist_count = RecordsetHelper.ToInt32(rs, "transit_pax_waitlist_count");
                        av.transit_waitlist_capacity = RecordsetHelper.ToInt32(rs, "transit_waitlist_capacity");
                        av.transit_close_hours = RecordsetHelper.ToInt32(rs, "transit_close_hours");
                        av.transit_reopen_hours = RecordsetHelper.ToInt32(rs, "transit_reopen_hours");
                        av.transit_nested_book_available = RecordsetHelper.ToInt32(rs, "transit_nested_book_available");
                        av.transit_physical_capacity = RecordsetHelper.ToInt32(rs, "transit_physical_capacity");
                        av.transit_bookable_capacity = RecordsetHelper.ToInt32(rs, "transit_bookable_capacity");
                        av.total_planned_departure_time = RecordsetHelper.ToInt32(rs, "total_planned_departure_time");
                        av.total_planned_arrival_time = RecordsetHelper.ToInt32(rs, "total_planned_arrival_time");
                        av.total_number_of_stops = RecordsetHelper.ToInt32(rs, "total_number_of_stops");
                        av.total_class_capacity = RecordsetHelper.ToInt32(rs, "total_class_capacity");
                        av.total_pax_book_count = RecordsetHelper.ToInt32(rs, "total_pax_book_count");
                        av.total_pax_waitlist_count = RecordsetHelper.ToInt32(rs, "total_pax_waitlist_count");
                        av.total_waitlist_capacity = RecordsetHelper.ToInt32(rs, "total_waitlist_capacity");
                        av.total_nested_book_available = RecordsetHelper.ToInt32(rs, "total_nested_book_available");
                        av.total_physical_capacity = RecordsetHelper.ToInt32(rs, "total_physical_capacity");
                        av.total_bookable_capacity = RecordsetHelper.ToInt32(rs, "total_bookable_capacity");
                        av.fare_column = RecordsetHelper.ToInt32(rs, "fare_column");
                        av.timelimit_hours_before_flight = RecordsetHelper.ToInt32(rs, "timelimit_hours_before_flight");
                        av.timelimit_hours_after_booking = RecordsetHelper.ToInt32(rs, "timelimit_hours_after_booking");

                        av.redemption_points = RecordsetHelper.ToDecimal(rs, "redemption_points");
                        av.adult_fare = RecordsetHelper.ToDecimal(rs, "adult_fare");
                        av.child_fare = RecordsetHelper.ToDecimal(rs, "child_fare");
                        av.infant_fare = RecordsetHelper.ToDecimal(rs, "infant_fare");
                        av.other_fare = RecordsetHelper.ToDecimal(rs, "other_fare");
                        av.transit_redemption_points = RecordsetHelper.ToDecimal(rs, "transit_redemption_points");
                        av.transit_adult_fare = RecordsetHelper.ToDecimal(rs, "transit_adult_fare");
                        av.transit_child_fare = RecordsetHelper.ToDecimal(rs, "transit_child_fare");
                        av.transit_infant_fare = RecordsetHelper.ToDecimal(rs, "transit_infant_fare");
                        av.transit_other_fare = RecordsetHelper.ToDecimal(rs, "transit_other_fare");
                        av.total_redemption_points = RecordsetHelper.ToDecimal(rs, "total_redemption_points");
                        av.total_adult_fare = RecordsetHelper.ToDecimal(rs, "total_adult_fare");
                        av.total_child_fare = RecordsetHelper.ToDecimal(rs, "total_child_fare");
                        av.total_infant_fare = RecordsetHelper.ToDecimal(rs, "total_infant_fare");
                        av.total_other_fare = RecordsetHelper.ToDecimal(rs, "total_other_fare");
                        av.adult_fare_excl = RecordsetHelper.ToDecimal(rs, "adult_fare_excl");
                        av.child_fare_excl = RecordsetHelper.ToDecimal(rs, "child_fare_excl");
                        av.infant_fare_excl = RecordsetHelper.ToDecimal(rs, "infant_fare_excl");
                        av.other_fare_excl = RecordsetHelper.ToDecimal(rs, "other_fare_excl");

                        availability.Add(av);

                        rs.MoveNext();
                    }
                }
            }
        }
        public static IList<Availability> RecordsetToAvailability(this ADODB.Recordset rs, string avalabilityLeg)
        {
            if (rs != null && rs.RecordCount > 0)
            {
                
                IList<Availability> objAvailability = new List<Availability>();
                Availability  av = null;

                rs.MoveFirst();
                while (!rs.EOF)
                {
                    av = new Availability();
                    
                    av.AvalabilityLeg = avalabilityLeg;
                    av.flight_id = RecordsetHelper.ToGuid(rs, "flight_id");
                    av.transit_fare_id = RecordsetHelper.ToGuid(rs, "transit_fare_id");
                    av.transit_flight_id = RecordsetHelper.ToGuid(rs, "transit_flight_id");
                    av.fare_id = RecordsetHelper.ToGuid(rs, "fare_id");

                    av.airline_rcd = RecordsetHelper.ToString(rs, "airline_rcd");
                    av.flight_number = RecordsetHelper.ToString(rs, "flight_number");
                    av.flight_status_rcd = RecordsetHelper.ToString(rs, "flight_status_rcd");
                    av.aircraft_type_rcd = RecordsetHelper.ToString(rs, "aircraft_type_rcd");
                    av.origin_rcd = RecordsetHelper.ToString(rs, "origin_rcd");
                    av.destination_rcd = RecordsetHelper.ToString(rs, "destination_rcd");
                    av.flight_comment = RecordsetHelper.ToString(rs, "flight_comment");
                    av.operating_airline_rcd = RecordsetHelper.ToString(rs, "operating_airline_rcd");
                    av.operating_flight_number = RecordsetHelper.ToString(rs, "operating_flight_number");
                    av.operating_airline_name = RecordsetHelper.ToString(rs, "operating_airline_name");
                    av.flight_information_1 = RecordsetHelper.ToString(rs, "flight_information_1");
                    av.flight_information_2 = RecordsetHelper.ToString(rs, "flight_information_2");
                    av.flight_information_3 = RecordsetHelper.ToString(rs, "flight_information_3");
                    av.endorsement_text = RecordsetHelper.ToString(rs, "endorsement_text");
                    av.restriction_text = RecordsetHelper.ToString(rs, "restriction_text");
                    av.origin_name = RecordsetHelper.ToString(rs, "origin_name");
                    av.destination_name = RecordsetHelper.ToString(rs, "destination_name");
                    av.flight_check_in_status_rcd = RecordsetHelper.ToString(rs, "flight_check_in_status_rcd");
                    av.fare_code = RecordsetHelper.ToString(rs, "fare_code");
                    av.fare_type_rcd = RecordsetHelper.ToString(rs, "fare_type_rcd");
                    av.booking_class_rcd = RecordsetHelper.ToString(rs, "booking_class_rcd"); 
                    av.boarding_class_rcd = RecordsetHelper.ToString(rs, "boarding_class_rcd"); 
                    av.nesting_string = RecordsetHelper.ToString(rs, "nesting_string"); 
                    av.origin_country_rcd = RecordsetHelper.ToString(rs, "origin_country_rcd"); 
                    av.destination_country_rcd = RecordsetHelper.ToString(rs, "destination_country_rcd"); 
                    av.transit_country_rcd = RecordsetHelper.ToString(rs, "transit_country_rcd"); 
                    av.transit_points = RecordsetHelper.ToString(rs, "transit_points"); 
                    av.transit_points_name = RecordsetHelper.ToString(rs, "transit_points_name"); 
                    av.transit_airport_rcd = RecordsetHelper.ToString(rs, "transit_airport_rcd"); 
                    av.transit_name = RecordsetHelper.ToString(rs, "transit_name"); 
                    av.transit_airline_rcd = RecordsetHelper.ToString(rs, "transit_airline_rcd"); 
                    av.transit_flight_number = RecordsetHelper.ToString(rs, "transit_flight_number"); 
                    av.transit_operating_airline_rcd = RecordsetHelper.ToString(rs, "transit_operating_airline_rcd"); 
                    av.transit_operating_flight_number = RecordsetHelper.ToString(rs, "transit_operating_flight_number"); 
                    av.transit_flight_status_rcd = RecordsetHelper.ToString(rs, "transit_flight_status_rcd"); 
                    av.transit_aircraft_type_rcd = RecordsetHelper.ToString(rs, "transit_aircraft_type_rcd"); 
                    av.transit_flight_comment = RecordsetHelper.ToString(rs, "transit_flight_comment"); 
                    av.transit_flight_check_in_status_rcd = RecordsetHelper.ToString(rs, "transit_flight_check_in_status_rcd"); 
                    av.transit_fare_code = RecordsetHelper.ToString(rs, "transit_fare_code"); 
                    av.transit_fare_type_rcd = RecordsetHelper.ToString(rs, "transit_fare_type_rcd"); 
                    av.transit_booking_class_rcd = RecordsetHelper.ToString(rs, "transit_booking_class_rcd"); 
                    av.transit_boarding_class_rcd = RecordsetHelper.ToString(rs, "transit_boarding_class_rcd"); 
                    av.currency_rcd = RecordsetHelper.ToString(rs, "currency_rcd"); 
                    av.promotion_code = RecordsetHelper.ToString(rs, "promotion_code"); 
                    av.promotion_text = RecordsetHelper.ToString(rs, "promotion_text"); 
                    av.transit_transit_points = RecordsetHelper.ToString(rs, "transit_transit_points"); 
                    av.transit_transit_points_name = RecordsetHelper.ToString(rs, "transit_transit_points_name"); 
                    av.transit_nesting_string = RecordsetHelper.ToString(rs, "transit_nesting_string"); 
                   // av.promotion_rcd = RecordsetHelper.ToString(rs, "promotion_rcd"); 
                    av.transaction_reference = RecordsetHelper.ToString(rs, "transaction_reference"); 

                    av.total_departure_date = RecordsetHelper.ToDateTime(rs, "total_departure_date"); 
                    av.departure_date = RecordsetHelper.ToDateTime(rs, "departure_date");
                    av.arrival_date = RecordsetHelper.ToDateTime(rs, "arrival_date");
                    av.transit_departure_date = RecordsetHelper.ToDateTime(rs, "transit_departure_date");
                    av.transit_arrival_date = RecordsetHelper.ToDateTime(rs, "transit_arrival_date");
                    av.utc_departure_date_time = RecordsetHelper.ToDateTime(rs, "utc_departure_date_time");
                    av.utc_arrival_date_time = RecordsetHelper.ToDateTime(rs, "utc_arrival_date_time"); 

                    av.planned_departure_time = RecordsetHelper.ToInt16(rs, "planned_departure_time");
                    av.estimated_departure_time = RecordsetHelper.ToInt16(rs, "estimated_departure_time");
                    av.planned_arrival_time = RecordsetHelper.ToInt16(rs, "planned_arrival_time");
                    av.flight_duration = RecordsetHelper.ToInt16(rs, "flight_duration");
                    av.number_of_stops = RecordsetHelper.ToInt16(rs, "number_of_stops");
                    av.departed = RecordsetHelper.ToInt16(rs, "departed");
                    av.transit_planned_departure_time = RecordsetHelper.ToInt16(rs, "transit_planned_departure_time");
                    av.transit_estimated_departure_time = RecordsetHelper.ToInt16(rs, "transit_estimated_departure_time");
                    av.transit_planned_arrival_time = RecordsetHelper.ToInt16(rs, "transit_planned_arrival_time"); 
                    av.time_diff = RecordsetHelper.ToInt32(rs, "time_diff"); 
                    av.close_web_sales = RecordsetHelper.ToInt32(rs, "close_web_sales"); 
                    av.total_flight_duration = RecordsetHelper.ToInt32(rs, "total_flight_duration"); 

                    av.it_fare_flag = RecordsetHelper.ToByte(rs, "it_fare_flag");
                    av.class_open_flag = RecordsetHelper.ToByte(rs, "class_open_flag");
                    av.waitlist_open_flag = RecordsetHelper.ToByte(rs, "waitlist_open_flag");
                    av.ignore_logic_flag = RecordsetHelper.ToByte(rs, "ignore_logic_flag");
                    av.transit_it_fare_flag = RecordsetHelper.ToByte(rs, "transit_it_fare_flag");
                    av.transit_class_open_flag = RecordsetHelper.ToByte(rs, "transit_class_open_flag");
                    av.transit_waitlist_open_flag = RecordsetHelper.ToByte(rs, "transit_waitlist_open_flag");
                    av.transit_ignore_logic_flag = RecordsetHelper.ToByte(rs, "transit_ignore_logic_flag");
                    av.total_class_open_flag = RecordsetHelper.ToByte(rs, "total_class_open_flag");
                    av.total_waitlist_open_flag = RecordsetHelper.ToByte(rs, "total_waitlist_open_flag");
                    av.corporate_fare_flag = RecordsetHelper.ToByte(rs, "corporate_fare_flag");
                    av.refundable_flag = RecordsetHelper.ToByte(rs, "refundable_flag");
                    av.full_flight_flag = RecordsetHelper.ToByte(rs, "full_flight_flag");
                    av.no_waitlist_flight_flag = RecordsetHelper.ToByte(rs, "no_waitlist_flight_flag");
                    av.departed_flight_flag = RecordsetHelper.ToByte(rs, "departed_flight_flag");
                    av.public_fare_flag = RecordsetHelper.ToByte(rs, "public_fare_flag");
                    av.open_return_fare_flag = RecordsetHelper.ToByte(rs, "open_return_fare_flag");
                    av.staff_fare_flag = RecordsetHelper.ToByte(rs, "staff_fare_flag");
                    av.subload_fare_flag = RecordsetHelper.ToByte(rs, "subload_fare_flag");
                    av.e_ticket_flag = RecordsetHelper.ToByte(rs, "e_ticket_flag");
                    av.avl_show_net_total_flag = RecordsetHelper.ToByte(rs, "avl_show_net_total_flag");
                    av.filter_logic_flag = RecordsetHelper.ToByte(rs, "filter_logic_flag"); 

                    av.class_capacity = RecordsetHelper.ToInt32(rs, "class_capacity");
                    av.pax_book_count = RecordsetHelper.ToInt32(rs, "pax_book_count");
                    av.pax_waitlist_count = RecordsetHelper.ToInt32(rs, "pax_waitlist_count");
                    av.waitlist_capacity = RecordsetHelper.ToInt32(rs, "waitlist_capacity");
                    av.close_hours = RecordsetHelper.ToInt32(rs, "close_hours");
                    av.reopen_hours = RecordsetHelper.ToInt32(rs, "reopen_hours");
                    av.nested_book_available = RecordsetHelper.ToInt32(rs, "nested_book_available");
                    av.physical_capacity = RecordsetHelper.ToInt32(rs, "physical_capacity");
                    av.bookable_capacity = RecordsetHelper.ToInt32(rs, "bookable_capacity");
                    av.transit_flight_duration = RecordsetHelper.ToInt32(rs, "transit_flight_duration");
                    av.transit_number_of_stops = RecordsetHelper.ToInt32(rs, "transit_number_of_stops");
                    av.transit_departed = RecordsetHelper.ToInt32(rs, "transit_departed");
                    av.transit_class_capacity = RecordsetHelper.ToInt32(rs, "transit_class_capacity");
                    av.transit_pax_book_count = RecordsetHelper.ToInt32(rs, "transit_pax_book_count");
                    av.transit_pax_waitlist_count = RecordsetHelper.ToInt32(rs, "transit_pax_waitlist_count");
                    av.transit_waitlist_capacity = RecordsetHelper.ToInt32(rs, "transit_waitlist_capacity");
                    av.transit_close_hours = RecordsetHelper.ToInt32(rs, "transit_close_hours");
                    av.transit_reopen_hours = RecordsetHelper.ToInt32(rs, "transit_reopen_hours");
                    av.transit_nested_book_available = RecordsetHelper.ToInt32(rs, "transit_nested_book_available");
                    av.transit_physical_capacity = RecordsetHelper.ToInt32(rs, "transit_physical_capacity");
                    av.transit_bookable_capacity = RecordsetHelper.ToInt32(rs, "transit_bookable_capacity");
                    av.total_planned_departure_time = RecordsetHelper.ToInt32(rs, "total_planned_departure_time");
                    av.total_planned_arrival_time = RecordsetHelper.ToInt32(rs, "total_planned_arrival_time");
                    av.total_number_of_stops = RecordsetHelper.ToInt32(rs, "total_number_of_stops");
                    av.total_class_capacity = RecordsetHelper.ToInt32(rs, "total_class_capacity");
                    av.total_pax_book_count = RecordsetHelper.ToInt32(rs, "total_pax_book_count");
                    av.total_pax_waitlist_count = RecordsetHelper.ToInt32(rs, "total_pax_waitlist_count");
                    av.total_waitlist_capacity = RecordsetHelper.ToInt32(rs, "total_waitlist_capacity");
                    av.total_nested_book_available = RecordsetHelper.ToInt32(rs, "total_nested_book_available");
                    av.total_physical_capacity = RecordsetHelper.ToInt32(rs, "total_physical_capacity");
                    av.total_bookable_capacity = RecordsetHelper.ToInt32(rs, "total_bookable_capacity");
                    av.fare_column = RecordsetHelper.ToInt32(rs, "fare_column");
                    av.timelimit_hours_before_flight = RecordsetHelper.ToInt32(rs, "timelimit_hours_before_flight");
                    av.timelimit_hours_after_booking = RecordsetHelper.ToInt32(rs, "timelimit_hours_after_booking"); 

                    av.redemption_points = RecordsetHelper.ToDecimal(rs, "redemption_points");
                    av.adult_fare = RecordsetHelper.ToDecimal(rs, "adult_fare");
                    av.child_fare = RecordsetHelper.ToDecimal(rs, "child_fare");
                    av.infant_fare = RecordsetHelper.ToDecimal(rs, "infant_fare");
                    av.other_fare = RecordsetHelper.ToDecimal(rs, "other_fare");
                    av.transit_redemption_points = RecordsetHelper.ToDecimal(rs, "transit_redemption_points");
                    av.transit_adult_fare = RecordsetHelper.ToDecimal(rs, "transit_adult_fare");
                    av.transit_child_fare = RecordsetHelper.ToDecimal(rs, "transit_child_fare");
                    av.transit_infant_fare = RecordsetHelper.ToDecimal(rs, "transit_infant_fare");
                    av.transit_other_fare = RecordsetHelper.ToDecimal(rs, "transit_other_fare");
                    av.total_redemption_points = RecordsetHelper.ToDecimal(rs, "total_redemption_points");
                    av.total_adult_fare = RecordsetHelper.ToDecimal(rs, "total_adult_fare");
                    av.total_child_fare = RecordsetHelper.ToDecimal(rs, "total_child_fare");
                    av.total_infant_fare = RecordsetHelper.ToDecimal(rs, "total_infant_fare");
                    av.total_other_fare = RecordsetHelper.ToDecimal(rs, "total_other_fare");
                    av.adult_fare_excl = RecordsetHelper.ToDecimal(rs, "adult_fare_excl");
                    av.child_fare_excl = RecordsetHelper.ToDecimal(rs, "child_fare_excl");
                    av.infant_fare_excl = RecordsetHelper.ToDecimal(rs, "infant_fare_excl");
                    av.other_fare_excl = RecordsetHelper.ToDecimal(rs, "other_fare_excl"); 

                    objAvailability.Add(av);
                }
                return objAvailability;
            }

            return null;
        }

        #endregion

    }
}
