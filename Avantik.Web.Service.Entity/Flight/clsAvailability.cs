using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

namespace Avantik.Web.Service.Entity.Flight
{
    public class Availability : FlightBase
    {
        #region Property
        public Guid flight_id { get; set; }
        public Guid transit_fare_id { get; set; }
        public Guid transit_flight_id { get; set; }
        public Guid fare_id { get; set; }

        public string AvalabilityLeg { get; set; }
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
       // public string promotion_rcd { get; set; }
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

        public byte transferable_fare_flag { get; set; }

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
        #endregion

        #region Method
        public void SetTotalTransitClassCapacity()
        {
            try
            {
                if (class_capacity - pax_book_count < transit_class_capacity - transit_pax_book_count)
                {
                    total_class_capacity = class_capacity;
                    total_pax_book_count = pax_book_count;
                }
                else
                {
                    total_class_capacity = transit_class_capacity;
                    total_pax_book_count = transit_pax_book_count;
                }
            }
            catch (Exception ex)
            {
                throw;
            }
            
        }
        public void SetTotalTransitWaitlistCapacity()
        {
            try
            {
                if (waitlist_capacity - pax_waitlist_count < transit_waitlist_capacity - transit_pax_waitlist_count)
                {
                    total_waitlist_capacity = waitlist_capacity;
                    total_pax_waitlist_count = pax_waitlist_count;
                }
                else
                {
                    total_waitlist_capacity = transit_waitlist_capacity;
                    total_pax_waitlist_count = transit_pax_waitlist_count;
                }
            }
            catch (Exception ex)
            {
                throw;
            }
        }
        public void SetTotalClassOpenFlag()
        {
            try
            {
                if (transit_flight_id.Equals(Guid.Empty))
                {
                    total_class_open_flag = class_open_flag;
                }
                else if (class_open_flag != 0 && transit_class_open_flag != 0)
                {
                    total_class_open_flag = 1;
                }
                else
                {
                    total_class_open_flag = 0;
                }
            }
            catch (Exception ex)
            {
                throw;
            }
        }
        public void SetTotalWailistOpenFlag()
        {
            try
            {
                if (transit_flight_id.Equals(Guid.Empty))
                {
                    total_waitlist_open_flag = waitlist_open_flag;
                }
                else if (waitlist_open_flag != 0 && transit_waitlist_open_flag != 0)
                {
                    total_waitlist_open_flag = 1;
                }
                else
                {
                    total_waitlist_open_flag = 0;
                }
            }
            catch (Exception ex)
            {
                throw;
            }
        }
        public void SetTotalNestBookAvai()
        {
            try
            {
                if (transit_flight_id.Equals(Guid.Empty))
                {
                    total_nested_book_available = nested_book_available;
                }
                else if (nested_book_available < transit_nested_book_available)
                {
                    total_nested_book_available = nested_book_available;
                }
                else
                {
                    total_nested_book_available = transit_nested_book_available;
                }
            }
            catch (Exception ex)
            {
                throw;
            }
        }
        public void SetTotalPhysicalCapacity()
        {
            try
            {
                if (transit_flight_id.Equals(Guid.Empty))
                {
                    total_physical_capacity = physical_capacity;
                }
                else if (physical_capacity < transit_physical_capacity)
                {
                    total_physical_capacity = physical_capacity;
                }
                else
                {
                    total_physical_capacity = transit_physical_capacity;
                }
            }
            catch (Exception ex)
            {
                throw;
            }
        }
        public void SetTotalBookableCapacity()
        {
            try
            {
                if (transit_flight_id.Equals(Guid.Empty))
                {
                    total_bookable_capacity = bookable_capacity;
                }
                else if (bookable_capacity < bookable_capacity)
                {
                    total_bookable_capacity = bookable_capacity;
                }
                else
                {
                    total_bookable_capacity = transit_bookable_capacity;
                }
            }
            catch (Exception ex)
            {
                throw;
            }
        }
        public void SetFullFlightFlag(byte adult, byte child, byte other)
        {
            try
            {
                if (adult + child + other <= total_nested_book_available)
                {
                    full_flight_flag = 0;
                }
                else
                {
                    full_flight_flag = 1;
                }
            }
            catch (Exception ex)
            {
                throw;
            }
        }
        public void SetTotalFare()
        {
            try
            {
                total_adult_fare = adult_fare + transit_adult_fare;
                total_child_fare = child_fare + transit_child_fare;
                total_infant_fare = infant_fare + transit_infant_fare;
                total_other_fare = other_fare + transit_other_fare;
            }
            catch (Exception ex)
            {
                throw;
            }
        }
        public void SetVAT(bool noVat)
        {
            if (noVat == true)
            {
                adult_fare = adult_fare_excl;
                child_fare = child_fare_excl;
                infant_fare = infant_fare_excl;
                other_fare = other_fare_excl;
            }
        }
        public bool FilterOutAvailability(byte adult,
                                    byte child,
                                    byte other,
                                    bool showClose)
        {
            //Check available fligt
            bool bCannotBook = CannotBook(flight_check_in_status_rcd, flight_status_rcd, estimated_departure_time, planned_arrival_time);
            bool bWaitlist = false;
            bool bFilterout = false;

            if (bCannotBook == false)
            {
                bWaitlist = Waitlist(adult, 
                                    child, 
                                    other, 
                                    class_open_flag, 
                                    waitlist_open_flag, 
                                    nested_book_available);

                //Filter out unused flight.
                bFilterout = FilterOutAvailability(showClose, 
                                                    bCannotBook, 
                                                    bWaitlist, 
                                                    class_open_flag, 
                                                    nested_book_available);

                if (bFilterout == false)
                {
                    //Check available flight for transit flight.
                    if (transit_flight_id.Equals(Guid.Empty) == false)
                    {
                        bCannotBook = CannotBook(transit_flight_check_in_status_rcd, 
                                                transit_flight_status_rcd, 
                                                transit_estimated_departure_time, 
                                                transit_planned_arrival_time);

                        if (bCannotBook == false)
                        {

                            bWaitlist = Waitlist(adult, 
                                                child, 
                                                other, 
                                                transit_class_open_flag, 
                                                transit_waitlist_open_flag, 
                                                transit_nested_book_available);

                            //Filter out unused transit flight.
                            bFilterout = FilterOutAvailability(showClose, 
                                                                bCannotBook, 
                                                                bWaitlist, 
                                                                transit_class_open_flag, 
                                                                transit_nested_book_available);
                        }
                    }
                }
            }

            return bFilterout;
        }

        #endregion

        #region Helper
        private bool CannotBook(string flightCheckinStatus,
                                string flightStatus,
                                short estimateDepartureTime,
                                short planedArrivalTime)
        {
            try
            {
                //check if flight is available to book.
                if (flightCheckinStatus == "OPEN" && estimateDepartureTime > planedArrivalTime)
                {
                    return false;
                }
                else if (string.IsNullOrEmpty(flightStatus) == true)
                {
                    return true;
                }
                else if (flightStatus == "INACTIVE")
                {
                    return true;
                }
                else if (flightStatus == "CANCELLED")
                {
                    return true;
                }
                else if (flightCheckinStatus == "CLOSED")
                {
                    return true;
                }
                else if (flightCheckinStatus == "DISPATCHED")
                {
                    return true;
                }
                else if (flightCheckinStatus == "FLOWN")
                {
                    return true;
                }

                return false;
            }
            catch (Exception ex)
            {
                throw;
            }

        }

        private bool Waitlist(byte adult,
                              byte child,
                              byte other,
                              byte classOpenFlag,
                              byte waitlistOpenFlag,
                              int nestedBookAvailable)
        {
            if (classOpenFlag == 0)
            {
                if (waitlistOpenFlag == 0)
                {
                    return false;
                }
                else
                {
                    return true;
                }
            }
            else
            {
                if (nestedBookAvailable < (adult + child + other))
                {
                    if (waitlistOpenFlag == 0)
                    {
                        return false;
                    }
                    else
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        private bool FilterOutAvailability(bool allowedShowClose,
                                           bool cannotBook,
                                           bool waitlist,
                                           byte classOpenFlag,
                                           int nestedBookAvailable)
        {
            if (allowedShowClose == false && classOpenFlag == 0)
            {
                return true;
            }
            else if (allowedShowClose == false && nestedBookAvailable == 0 && waitlist == false)
            {
                return true;
            }
            else if (cannotBook == true && allowedShowClose == false)
            {
                return true;
            }
            else if (allowedShowClose == false && classOpenFlag == 0)
            {
                return true;
            }

            return false;
        }
        #endregion
    }
}
