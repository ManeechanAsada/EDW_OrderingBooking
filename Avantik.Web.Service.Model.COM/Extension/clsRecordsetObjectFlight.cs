using Avantik.Web.Service.Entity.Booking;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Avantik.Web.Service.COMHelper;
using Avantik.Web.Service.Exception.Booking;
using entity = Avantik.Web.Service.Entity;

namespace Avantik.Web.Service.Model.COM.Extension
{
    public static class RecordsetObjectFlight
    {
        public static void ToRecordset(this  IList<entity.Booking.Flight> flights, ref ADODB.Recordset rs)
        {
            if (flights != null && flights.Count > 0)
            {
                try
                {
                    foreach (entity.Booking.Flight fs in flights)
                    {
                        rs.AddNew();
                        rs.Fields["boarding_class_rcd"].Value = fs.BoardingClassRcd;
                        rs.Fields["booking_class_rcd"].Value = fs.BookingClassRcd;

                        if (fs.FlightId != Guid.Empty)
                            rs.Fields["flight_id"].Value = fs.FlightId.ToRsString();

                        rs.Fields["origin_rcd"].Value = fs.OriginRcd;
                        rs.Fields["departure_date"].Value = fs.DepartureDate;

                        if (fs.FairId != Guid.Empty)
                            rs.Fields["fare_id"].Value = fs.FairId.ToRsString();

                        if (fs.ExchangedSegmentId != Guid.Empty)
                            rs.Fields["exchanged_segment_id"].Value = fs.ExchangedSegmentId.ToRsString();

                        rs.Fields["eticket_flag"].Value = fs.EticketFlag;
                        rs.Fields["od_origin_rcd"].Value = fs.OdOriginRcd;
                        rs.Fields["od_destination_rcd"].Value = fs.OdDestinationRcd;

                        if (fs.FlightConnectionId != Guid.Empty)
                            rs.Fields["flight_connection_id"].Value = fs.FlightConnectionId.ToRsString();

                        rs.Fields["destination_rcd"].Value = fs.DestinationRcd;
                        rs.Fields["transit_points"].Value = fs.TransitPoints;

                        rs.Fields["airline_rcd"].Value = fs.AirlineRcd;
                        rs.Fields["flight_number"].Value = fs.FlightNumber;
                        rs.Fields["number_of_units"].Value = fs.NumberOfUnits;
                    }

                }
                catch(Exception.Booking.BookingException ex)
                {
                    throw ex;
                }

            }
        }
    }
}
