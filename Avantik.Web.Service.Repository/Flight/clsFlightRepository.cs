using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;
using Avantik.Web.Service.Entity.Flight;
using Avantik.Web.Service.Helpers;
using Avantik.Web.Service.Helpers.Database;
using Avantik.Web.Service.Repository.Contract.Flight;

namespace Avantik.Web.Service.Repository.Flight
{
    public class FlightRepository : IFlightRepository
    {
        string _connectionString;
        public FlightRepository() { 
        
        }
        public FlightRepository(string connectionString)
        {
            _connectionString = connectionString;
        }
        public IEnumerable<AvailabilityRoute> GetFlightAvailabilityRoute(string originRcd, string destinationRcd)
        {

            try
            {
                IList<AvailabilityRoute> avaiRoute = null;
                SqlDataReader rw;
                //Open DB connection.
                using (DBHelper db = new DBHelper(_connectionString))
                {
                    db.Connect();
                    //Call store procedure.
                    rw = db.ExecDataReaderProc("get_flight_availability_routes",
                                                                                "@origin", originRcd,
                                                                                "@destination", destinationRcd);

                    if (rw != null && rw.HasRows == true)
                    {
                        avaiRoute = new List<AvailabilityRoute>();
                        AvailabilityRoute route = null;
                        while (rw.Read())
                        {
                            route = new AvailabilityRoute();
                            route.origin_rcd = rw.DBToString("origin_rcd");
                            route.destination_rcd = rw.DBToString("destination_rcd");
                            route.segment_change_fee_flag = rw.DBToByte("segment_change_fee_flag");
                            route.special_service_fee_flag = rw.DBToByte("special_service_fee_flag");
                            route.show_insurance_on_web_flag = rw.DBToByte("show_insurance_on_web_flag");
                            route.transit_airport_rcd = rw.DBToString("transit_airport_rcd");
                            route.avl_origin_rcd = rw.DBToString("avl_origin_rcd");
                            route.avl_destination_rcd = rw.DBToString("avl_destination_rcd");
                            route.transit_flag = rw.DBToByte("transit_flag");
                            route.direct_flag = rw.DBToByte("direct_flag");
                            route.dynamic_connections_flag = rw.DBToByte("dynamic_connections_flag");
                            route.min_transit_minutes = rw.DBToInt32("min_transit_minutes");
                            route.max_transit_minutes = rw.DBToInt32("max_transit_minutes");
                            route.connex = rw.DBToInt32("connex");

                            avaiRoute.Add(route);
                            route = null;
                        }
                        
                    }
                }
                
                return avaiRoute;
            }
            catch (Exception ex)
            {
                throw;
            }
        }
        public IEnumerable<AvailabilityParing> GetFlightAvailabilityParing(string originRcd,
                                                                            string destinationRcd,
                                                                            string transit,
                                                                            DateTime dateFrom,
                                                                            DateTime dateTo)
        {
            try
            {
                IList<AvailabilityParing> avaiParing = null;
                SqlDataReader rw;
                //Open DB connection.
                using (DBHelper db = new DBHelper(_connectionString))
                {
                    db.Connect();
                    //Call store procedure.
                    rw = db.ExecDataReaderProc("get_flight_availability_pairing",
                                                                                "@origin", originRcd,
                                                                                "@destination", destinationRcd,
                                                                                "@transit", transit,
                                                                                "@datefrom", dateFrom.GetDateString(),
                                                                                "@dateto", dateTo.GetDateString());
                    if (rw != null && rw.HasRows == true)
                    {
                        avaiParing = new List<AvailabilityParing>();
                        AvailabilityParing paring = null;
                        while (rw.Read())
                        {
                            paring = new AvailabilityParing();

                            paring.origin_rcd = rw.DBToString("origin_rcd");
                            paring.destination_rcd = rw.DBToString("destination_rcd");
                            paring.transit_airport_rcd = rw.DBToString("transit_airport_rcd");
                            paring.transit_flag = rw.DBToByte("transit_flag");
                            paring.direct_flag = rw.DBToByte("direct_flag");
                            paring.dynamic_connections_flag = rw.DBToByte("dynamic_connections_flag");
                            paring.min_transit_minutes = rw.DBToInt32("min_transit_minutes");
                            paring.max_transit_minutes = rw.DBToInt32("max_transit_minutes");
                            paring.leg_1_airline_rcd = rw.DBToString("leg_1_airline_rcd");
                            paring.leg_1_flight_number = rw.DBToString("leg_1_flight_number");
                            paring.leg_2_airline_rcd = rw.DBToString("leg_2_airline_rcd");
                            paring.leg_2_flight_number = rw.DBToString("leg_2_flight_number");

                            avaiParing.Add(paring);
                            paring = null;
                        }
                    }
                }

                return avaiParing;
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        public IEnumerable<AvailabilityQuoteTax> GetFlightAvailabilityQuoteTax(string originRcd,
                                                                                string destinationRcd,
                                                                                string currencyRcd,
                                                                                string passengerTypeRcd,
                                                                                string agencyCode,
                                                                                DateTime flightDate,
                                                                                DateTime bookingDate)
        {
            try
            {
                IList<AvailabilityQuoteTax> quoteTaxes = null;
                SqlDataReader rw;
                //Open DB connection.
                using (DBHelper db = new DBHelper(_connectionString))
                {
                    db.Connect();
                    //Call store procedure.
                    rw = db.ExecDataReaderProc("get_fare_quote_tax_avl",
                                                                    "@origin", originRcd,
                                                                    "@destination", destinationRcd,
                                                                    "@currency", currencyRcd,
                                                                    "@paxtype", passengerTypeRcd,
                                                                    "@flightdate", flightDate.GetDateString(),
                                                                    "@bookingdate", bookingDate.GetDateString(),
                                                                    "@agencycode", agencyCode);

                    if (rw != null && rw.HasRows == true)
                    {
                        quoteTaxes = new List<AvailabilityQuoteTax>();
                        AvailabilityQuoteTax quoteTax = null;
                        while (rw.Read())
                        {
                            quoteTax = new AvailabilityQuoteTax();

                            quoteTax.tax_rcd = rw.DBToString("tax_rcd");
                            quoteTax.airport_rcd = rw.DBToString("airport_rcd");
                            quoteTax.country_rcd = rw.DBToString("country_rcd");
                            quoteTax.valid_for_class = rw.DBToString("valid_for_class");
                            quoteTax.origin_tax_distribution = rw.DBToDecimal("origin_tax_distribution");
                            quoteTax.destination_tax_distribution = rw.DBToDecimal("destination_tax_distribution");
                            quoteTax.fare_code = rw.DBToString("fare_code");
                            quoteTax.tax_currency = rw.DBToString("tax_currency");

                            quoteTax.transit_tax_only_flag = rw.DBToByte("transit_tax_only_flag");
                            quoteTax.include_surcharge_flag = rw.DBToByte("include_surcharge_flag");
                            quoteTax.minimum_tax_amount_flag = rw.DBToByte("minimum_tax_amount_flag");
                            quoteTax.fare_basis_flag = rw.DBToByte("fare_basis_flag");

                            quoteTax.tax_amount = rw.DBToDecimal("tax_amount");
                            quoteTax.vat_percentage = rw.DBToDecimal("vat_percentage");
                            quoteTax.tax_percentage = rw.DBToDecimal("tax_percentage");
                            quoteTax.tax_amount_incl = rw.DBToDecimal("tax_amount_incl");
                            quoteTax.exchange_to_accounting = rw.DBToDecimal("exchange_to_accounting");
                            quoteTax.exchange_from_accounting = rw.DBToDecimal("exchange_from_accounting");

                            //Fill value from parameter.
                            quoteTax.passenger_type_rcd = passengerTypeRcd;
                            quoteTax.departure_date = flightDate;

                            quoteTaxes.Add(quoteTax);
                            quoteTax = null;
                        }
                    }
                }

                return quoteTaxes;
            }
            catch (Exception ex)
            {
                throw;
            }
        }
        public string GetFareLogicBookingClass()
        {
            return string.Empty;
        }
    }
}
