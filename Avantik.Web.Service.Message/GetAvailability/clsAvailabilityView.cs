using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

namespace Avantik.Web.Service.Message
{
    public class AvailabilityView
    {

        public Guid flight_id { get; set; }
        public Guid transit_fare_id { get; set; }
        public Guid transit_flight_id { get; set; }
        public Guid fare_id { get; set; }

        public string airline_rcd { get; set; }
        public string flight_number { get; set; }
        public string flight_status_rcd { get; set; }
        public string aircraft_type_rcd { get; set; }
        public string origin_rcd { get; set; }
        public string destination_rcd { get; set; }
        public string flight_comment { get; set; }
        public string operating_airline_rcd { get; set; }
        public string operating_flight_number { get; set; }
        public string operating_airline_name { get; set; }
        public string flight_information_1 { get; set; }
        public string flight_information_2 { get; set; }
        public string flight_information_3 { get; set; }
        public string endorsement_text { get; set; }
        public string restriction_text { get; set; }
        public string origin_name { get; set; }
        public string destination_name { get; set; }
        public string flight_check_in_status_rcd { get; set; }
        public string fare_code { get; set; }
        public string fare_type_rcd { get; set; }
        public string booking_class_rcd { get; set; }
        public string boarding_class_rcd { get; set; }
        public string nesting_string { get; set; }
        public string origin_country_rcd { get; set; }
        public string destination_country_rcd { get; set; }
        public string transit_country_rcd { get; set; }
        public string transit_points { get; set; }
        public string transit_points_name { get; set; }
        public string transit_airport_rcd { get; set; }
        public string transit_name { get; set; }
        public string transit_airline_rcd { get; set; }
        public string transit_flight_number { get; set; }
        public string transit_operating_airline_rcd { get; set; }
        public string transit_operating_flight_number { get; set; }
        public string transit_flight_status_rcd { get; set; }
        public string transit_aircraft_type_rcd { get; set; }
        public string transit_flight_comment { get; set; }
        public string transit_flight_check_in_status_rcd { get; set; }
        public string transit_fare_code { get; set; }
        public string transit_fare_type_rcd { get; set; }
        public string transit_booking_class_rcd { get; set; }
        public string transit_boarding_class_rcd { get; set; }
        public string currency_rcd { get; set; }
        public string promotion_code { get; set; }
        public string promotion_text { get; set; }
        public string transit_transit_points { get; set; }
        public string transit_transit_points_name { get; set; }
        public string transit_nesting_string { get; set; }
        public string promotion_rcd { get; set; }
        public string transaction_reference { get; set; }

        public DateTime total_departure_date { get; set; }
        public DateTime departure_date { get; set; }
        public DateTime arrival_date { get; set; }
        public DateTime transit_departure_date { get; set; }
        public DateTime transit_arrival_date { get; set; }
        public DateTime utc_departure_date_time { get; set; }
        public DateTime utc_arrival_date_time { get; set; }

        public Int16 planned_departure_time { get; set; }
        public Int16 estimated_departure_time { get; set; }
        public Int16 planned_arrival_time { get; set; }
        public Int16 flight_duration { get; set; }
        public Int16 number_of_stops { get; set; }
        public Int16 departed { get; set; }
        public Int16 transit_planned_departure_time { get; set; }
        public Int16 transit_estimated_departure_time { get; set; }
        public Int16 transit_planned_arrival_time { get; set; }
        public Int64 time_diff { get; set; }
        public Int64 close_web_sales { get; set; }
        public Int64 total_flight_duration { get; set; }

        public byte it_fare_flag { get; set; }
        public byte class_open_flag { get; set; }
        public byte waitlist_open_flag { get; set; }
        public byte ignore_logic_flag { get; set; }
        public byte transit_it_fare_flag { get; set; }
        public byte transit_class_open_flag { get; set; }
        public byte transit_waitlist_open_flag { get; set; }
        public byte transit_ignore_logic_flag { get; set; }
        public byte total_class_open_flag { get; set; }
        public byte total_waitlist_open_flag { get; set; }
        public byte corporate_fare_flag { get; set; }
        public byte refundable_flag { get; set; }
        public byte full_flight_flag { get; set; }
        public byte no_waitlist_flight_flag { get; set; }
        public byte departed_flight_flag { get; set; }
        public byte public_fare_flag { get; set; }
        public byte open_return_fare_flag { get; set; }
        public byte staff_fare_flag { get; set; }
        public byte subload_fare_flag { get; set; }
        public byte e_ticket_flag { get; set; }
        public byte avl_show_net_total_flag { get; set; }

        public int filter_logic_flag { get; set; }
        public int class_capacity { get; set; }
        public int pax_book_count { get; set; }
        public int pax_waitlist_count { get; set; }
        public int waitlist_capacity { get; set; }
        public int close_hours { get; set; }
        public int reopen_hours { get; set; }
        public int nested_book_available { get; set; }
        public int physical_capacity { get; set; }
        public int bookable_capacity { get; set; }
        public int transit_flight_duration { get; set; }
        public int transit_number_of_stops { get; set; }
        public int transit_departed { get; set; }
        public int transit_class_capacity { get; set; }
        public int transit_pax_book_count { get; set; }
        public int transit_pax_waitlist_count { get; set; }
        public int transit_waitlist_capacity { get; set; }
        public int transit_close_hours { get; set; }
        public int transit_reopen_hours { get; set; }
        public int transit_nested_book_available { get; set; }
        public int transit_physical_capacity { get; set; }
        public int transit_bookable_capacity { get; set; }
        public int total_planned_departure_time { get; set; }
        public int total_planned_arrival_time { get; set; }
        public int total_number_of_stops { get; set; }
        public int total_class_capacity { get; set; }
        public int total_pax_book_count { get; set; }
        public int total_pax_waitlist_count { get; set; }
        public int total_waitlist_capacity { get; set; }
        public int total_nested_book_available { get; set; }
        public int total_physical_capacity { get; set; }
        public int total_bookable_capacity { get; set; }
        public int fare_column { get; set; }
        public int timelimit_hours_before_flight { get; set; }
        public int timelimit_hours_after_booking { get; set; }

        public decimal redemption_points { get; set; }
        public decimal adult_fare { get; set; }
        public decimal child_fare { get; set; }
        public decimal infant_fare { get; set; }
        public decimal other_fare { get; set; }
        public decimal transit_redemption_points { get; set; }
        public decimal transit_adult_fare { get; set; }
        public decimal transit_child_fare { get; set; }
        public decimal transit_infant_fare { get; set; }
        public decimal transit_other_fare { get; set; }
        public decimal total_redemption_points { get; set; }
        public decimal total_adult_fare { get; set; }
        public decimal total_child_fare { get; set; }
        public decimal total_infant_fare { get; set; }
        public decimal total_other_fare { get; set; }
        public decimal adult_fare_excl { get; set; }
        public decimal child_fare_excl { get; set; }
        public decimal infant_fare_excl { get; set; }
        public decimal other_fare_excl { get; set; }
    }
}
