using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;
using Avantik.Web.Service.Entity.Flight;
using Avantik.Web.Service.Helpers;
using Avantik.Web.Service.Helpers.Database;
using Avantik.Web.Service.Infrastructure;
using Avantik.Web.Service.Repository;

namespace Avantik.Web.Service.Repository.Flight
{
    public class AvailabilityRepository : Contract.Flight.IAvailabilityRepository
    {
        string _connectionString;
        public AvailabilityRepository() { 
        
        }
        public AvailabilityRepository(string connectionString)
        {
            _connectionString = connectionString;
        }

        public IEnumerable<Availability> FindAvailability(string originRcd,
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
                                                            string transactionReference)
        {

            try
            {
                if (string.IsNullOrEmpty(originRcd))
                {
                    throw new ArgumentException("Origin rcd is required.");
                }
                else if (string.IsNullOrEmpty(destinationRcd))
                {
                    throw new ArgumentException("Destination rcd is required.");
                }
                else if (dateFrom.Equals(DateTime.MinValue))
                {
                    throw new ArgumentException("Date from is required.");
                }
                else
                {
                    //Get store procedure name.
                    string spName = GetAvailabilitySpName(mapWithFare,
                                                          fareId,
                                                          maxFare,
                                                          departed,
                                                          cancelled,
                                                          groupFare,
                                                          itFares,
                                                          staffFares);

                    //Get Availability from route information.
                    IList<Availability> availabilities = null;
                    SqlDataReader rw;

                    using (DBHelper db = new DBHelper(_connectionString))
                    {
                        //Open DB connection.
                        db.Connect();
                        rw = db.ExecDataReaderProc(spName,
                                                          "@origin", originRcd,
                                                          "@destination", destinationRcd,
                                                          "@datefrom", dateFrom.GetDateString(),
                                                          "@dateto", dateTo.GetDateString(),
                                                          "@bookingdate", bookingDate.GetDateString(),
                                                          "@bookingclass", bookingClass,
                                                          "@boardingclass", boardingClass,
                                                          "@otherpax", otherPax,
                                                          "@agency", agencyCode,
                                                          "@currency", currencyCode,
                                                          "@flightid", flightId.SetGuid(),
                                                          "@fareid", fareId.SetGuid(),
                                                          "@daytime", dayTime,
                                                          "@adult", adult,
                                                          "@child", child,
                                                          "@infant", infant,
                                                          "@other", other,
                                                          "@nonstop", nonstop,
                                                          "@groupfares", groupFare,
                                                          "@refundable", refundable,
                                                          "@cancelled", cancelled,
                                                          "@departed", departed,
                                                          "@waitlisted", waitlist,
                                                          "@soldout", soldout,
                                                          "@itfares", itFares,
                                                          "@stafffares", staffFares,
                                                          "@maxfare", maxFare,
                                                          "@includefares", includeFares,
                                                          "@od_origin", odOriginRcd,
                                                          "@od_destination", odDestinationRcd,
                                                          "@promotion", promotionCode,
                                                          "@faretype", GetFareTypeText(fareType),
                                                          "@language", languageCode,
                                                          "@ipaddress", ipAddress);

                        

                        if (rw != null && rw.HasRows)
                        {
                            availabilities = new List<Availability>();
                            Availability a = null;
                            
                            while (rw.Read())
                            {
                                a = new Availability();

                                a.flight_id = rw.DBToGuid("flight_id");
                                a.airline_rcd = rw.DBToString("airline_rcd");
                                a.flight_number = rw.DBToString("flight_number");
                                a.operating_airline_rcd = rw.DBToString("operating_airline_rcd");
                                a.operating_flight_number = rw.DBToString("operating_flight_number");
                                a.operating_airline_name = rw.DBToString("operating_airline_name");
                                a.flight_status_rcd = rw.DBToString("flight_status_rcd");
                                a.aircraft_type_rcd = rw.DBToString("aircraft_type_rcd");
                                a.flight_comment = rw.DBToString("flight_comment");
                                a.flight_information_1 = rw.DBToString("flight_information_1");
                                a.flight_information_2 = rw.DBToString("flight_information_2");
                                a.flight_information_3 = rw.DBToString("flight_information_3");
                                a.endorsement_text = rw.DBToString("endorsement_text");
                                a.restriction_text = rw.DBToString("restriction_text");
                                a.origin_rcd = rw.DBToString("origin_rcd");
                                a.origin_name = rw.DBToString("origin_name");
                                a.destination_rcd = rw.DBToString("destination_rcd");
                                a.destination_name = rw.DBToString("destination_name");
                                a.flight_check_in_status_rcd = rw.DBToString("flight_check_in_status_rcd");
                                a.departure_date = rw.DBToDateTime("departure_date");
                                a.planned_departure_time = rw.DBToInt16("planned_departure_time");
                                a.estimated_departure_time = rw.DBToInt16("estimated_departure_time");
                                a.planned_arrival_time = rw.DBToInt16("planned_arrival_time");
                                a.arrival_date = rw.DBToDateTime("arrival_date");
                                a.flight_duration = rw.DBToInt16("flight_duration");
                                a.number_of_stops = rw.DBToInt16("number_of_stops");
                                a.departed = rw.DBToInt16("departed");
                                a.fare_code = rw.DBToString("fare_code");
                                a.fare_type_rcd = rw.DBToString("fare_type_rcd");
                                a.fare_id = rw.DBToGuid("fare_id");
                                a.it_fare_flag = rw.DBToByte("it_fare_flag");
                                a.booking_class_rcd = rw.DBToString("booking_class_rcd");
                                a.boarding_class_rcd = rw.DBToString("boarding_class_rcd");
                                a.class_open_flag = rw.DBToByte("class_open_flag");
                                a.class_capacity = rw.DBToInt32("class_capacity");
                                a.pax_book_count = rw.DBToInt32("pax_book_count");
                                a.pax_waitlist_count = rw.DBToInt32("pax_waitlist_count");
                                a.waitlist_capacity = rw.DBToInt32("waitlist_capacity");
                                a.waitlist_open_flag = rw.DBToByte("waitlist_open_flag");
                                a.nesting_string = rw.DBToString("nesting_string");
                                a.close_hours = rw.DBToInt32("close_hours");
                                a.reopen_hours = rw.DBToInt32("reopen_hours");
                                a.ignore_logic_flag = rw.DBToByte("ignore_logic_flag");
                                a.nested_book_available = rw.DBToInt32("nested_book_available");
                                a.physical_capacity = rw.DBToInt32("physical_capacity");
                                a.bookable_capacity = rw.DBToInt32("bookable_capacity");
                                a.redemption_points = Convert.ToDecimal(rw.DBToDecimal("redemption_points").ToString("F"));//rw.DBToDecimal("redemption_points");
                                a.adult_fare = Convert.ToDecimal(rw.DBToDecimal("adult_fare").ToString("F"));
                                a.child_fare = Convert.ToDecimal(rw.DBToDecimal("child_fare").ToString("F"));//rw.DBToDecimal("child_fare");
                                a.infant_fare = Convert.ToDecimal(rw.DBToDecimal("infant_fare").ToString("F"));//rw.DBToDecimal("infant_fare");
                                a.other_fare = Convert.ToDecimal(rw.DBToDecimal("other_fare").ToString("F"));//rw.DBToDecimal("other_fare");

                                a.adult_fare_excl = Convert.ToDecimal(rw.DBToDecimal("adult_fare_excl").ToString("F"));//rw.DBToDecimal("adult_fare_excl"); ;
                                a.child_fare_excl = Convert.ToDecimal(rw.DBToDecimal("child_fare_excl").ToString("F"));//rw.DBToDecimal("child_fare_excl"); ;
                                a.infant_fare_excl = Convert.ToDecimal(rw.DBToDecimal("infant_fare_excl").ToString("F"));//rw.DBToDecimal("infant_fare_excl"); ;
                                a.other_fare_excl = Convert.ToDecimal(rw.DBToDecimal("other_fare_excl").ToString("F"));//rw.DBToDecimal("other_fare_excl"); ;

                                a.origin_country_rcd = rw.DBToString("origin_country_rcd");
                                a.destination_country_rcd = rw.DBToString("destination_country_rcd");
                                
                                a.transit_country_rcd = rw.DBToString("transit_country_rcd");
                                a.transit_points = rw.DBToString("transit_points");
                                a.transit_points_name = rw.DBToString("transit_points_name");
                                
                                a.promotion_text = rw.DBToString("promotion_text");
                                a.corporate_fare_flag = rw.DBToByte("corporate_fare_flag");
                                a.refundable_flag = rw.DBToByte("refundable_flag");

                                a.time_diff = rw.DBToInt64("time_diff");
                                a.close_web_sales = rw.DBToInt64("close_web_sales");
                                a.open_return_fare_flag = rw.DBToByte("open_return_fare_flag");
                                a.staff_fare_flag = rw.DBToByte("staff_fare_flag");
                                a.subload_fare_flag = rw.DBToByte("subload_fare_flag");
                                a.fare_column = rw.DBToInt32("fare_column");
                                a.timelimit_hours_before_flight = rw.DBToInt32("timelimit_hours_before_flight");
                                a.timelimit_hours_after_booking = rw.DBToInt32("timelimit_hours_after_booking");
                                a.e_ticket_flag = rw.DBToByte("e_ticket_flag");
                                a.currency_rcd = rw.DBToString("currency_rcd");
                                a.utc_arrival_date_time = rw.DBToDateTime("utc_arrival_date_time");
                                a.utc_departure_date_time = rw.DBToDateTime("utc_departure_date_time");
                                a.avl_show_net_total_flag = rw.DBToByte("avl_show_net_total_flag");

                                availabilities.Add(a);
                                a = null;
                            }
                        }
                    }
                    return availabilities;
                }
            }
            catch (Exception ex)
            {
                throw;
            }            
        }

        

        #region Helper
        private string GetAvailabilitySpName(bool mapWithFare,
                                            Guid fareId,
                                            decimal maxAmount,
                                            Int16 includeDeparted,
                                            Int16 includeCancelled,
                                            Int16 groupFares,
                                            Int16 itFareOnly,
                                            Int16 staffFaresOnly)
        {
            if (mapWithFare == true)
            {
                if (fareId.Equals(Guid.Empty) &
                    maxAmount == 0 &
                    includeDeparted == 0 &
                    includeCancelled == 0 &
                    groupFares == 0 &
                    itFareOnly == 0 &
                    staffFaresOnly == 0)
                {
                    return "get_flight_availability_fare_basic";
                }
                else
                {
                    return "get_flight_availability_fare";
                }
            }
            else
            {
                return "get_flight_availability";
            }
        }
        private string GetFareTypeText(AvailabilityFareTypes fareType)
        {
            if (fareType == AvailabilityFareTypes.REDEMPTION)
            {
                return "POINT";
            }
            else
            {
                return "FARE";
            }
        }
        #endregion

    }
}
