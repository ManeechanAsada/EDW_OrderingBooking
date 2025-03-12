using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Avantik.Web.Service.Entity.Flight;
using Avantik.Web.Service.COMHelper;
using Avantik.Web.Service.Exception.Booking;

namespace Avantik.Web.Service.Model.COM.Extension
{
    public static class RecordsetObjectSeatMap
    {
        public static void FillSeatMap(this IList<SeatMap> seatMaps, ref ADODB.Recordset rs)
        {
            if (rs != null && rs.RecordCount > 0)
            {
                SeatMap seatMap = null;

                try
                {
                    rs.MoveFirst();
                    while (!rs.EOF)
                    {
                        seatMap = new SeatMap();
                        seatMap.AircraftConfigurationCode = RecordsetHelper.ToString(rs, "aircraft_configuration_code");
                        seatMap.AirlineRcd = RecordsetHelper.ToString(rs, "airline_rcd");
                        seatMap.AisleFlag = RecordsetHelper.ToInt32(rs, "aisle_flag");
                        seatMap.BassinetFlag = RecordsetHelper.ToInt32(rs, "bassinet_flag");
                        seatMap.BlockB2bFlag = RecordsetHelper.ToInt32(rs, "block_b2b_flag");
                        seatMap.BlockB2cFlag = RecordsetHelper.ToInt32(rs, "block_b2c_flag");
                        seatMap.BlockedFlag = RecordsetHelper.ToInt32(rs, "blocked_flag");
                        seatMap.BoardingClassRcd = RecordsetHelper.ToString(rs, "boarding_class_rcd");
                        seatMap.BookingClassRcd = RecordsetHelper.ToString(rs, "booking_class_rcd");
                        seatMap.DepartureDate = RecordsetHelper.ToDateTime(rs, "departure_date");
                        seatMap.DestinationRcd = RecordsetHelper.ToString(rs, "destination_rcd");
                        seatMap.EmergencyExitFlag = RecordsetHelper.ToInt32(rs, "emergency_exit_flag");
                        seatMap.EticketFlag = RecordsetHelper.ToByte(rs, "eticket_flag");
                        seatMap.FairId = RecordsetHelper.ToGuid(rs, "fair_id");
                        seatMap.FlightCheckInStatusRcd = RecordsetHelper.ToString(rs, "flight_checkin_status_rcd");
                        seatMap.FlightConnectionId = RecordsetHelper.ToGuid(rs,"flight_connection_id");
                        seatMap.FlightId = RecordsetHelper.ToGuid(rs, "flight_id");
                        seatMap.FlightNumber = RecordsetHelper.ToString(rs, "flight_number");
                        seatMap.FreeSeatingFlag = RecordsetHelper.ToInt32(rs, "free_seating_flag");
                        seatMap.HanddicappedFlag = RecordsetHelper.ToInt32(rs, "handdicapped_flag");
                        seatMap.InfantFlag = RecordsetHelper.ToInt32(rs, "infant_flag");
                        seatMap.LayoutColumn = RecordsetHelper.ToInt32(rs, "layout_column");
                        seatMap.LayoutRow = RecordsetHelper.ToInt32(rs, "layout_row");
                        seatMap.LocationTypeRcd = RecordsetHelper.ToString(rs, "location_type_rcd");
                        seatMap.LowComfortFlag = RecordsetHelper.ToInt32(rs, "low_comfort_flag");
                        seatMap.NoChildFlag = RecordsetHelper.ToInt32(rs, "no_child_flag");
                        seatMap.NoInfantFlag = RecordsetHelper.ToInt32(rs, "no_infant_flag");
                        seatMap.NumberOfBays = RecordsetHelper.ToInt32(rs, "number_of_bays");
                        seatMap.NumberOfColumns = RecordsetHelper.ToInt32(rs, "number_of_columns");
                        seatMap.NumberOfRows = RecordsetHelper.ToInt32(rs, "number_of_rows");
                        seatMap.OdDestinationRcd = RecordsetHelper.ToString(rs, "od_destination_rcd");
                        seatMap.OdOriginRcd = RecordsetHelper.ToString(rs, "od_origin_rcd");
                        seatMap.OriginRcd = RecordsetHelper.ToString(rs, "origin_rcd");
                        seatMap.PassengerCount = RecordsetHelper.ToInt32(rs, "passenger_count");
                        seatMap.SeatColumn = RecordsetHelper.ToString(rs, "seat_column");
                        seatMap.FeeRcd =RecordsetHelper.ToString(rs, "fee_rcd");
                        seatMap.SeatRow = RecordsetHelper.ToInt32(rs, "seat_row");
                        seatMap.StretcherFlag = RecordsetHelper.ToInt32(rs, "stretcher_flag");
                        seatMap.TransitPoints = RecordsetHelper.ToString(rs, "transit_points");
                        seatMap.TransitpointsName = RecordsetHelper.ToString(rs, "transit_points_name");
                        seatMap.UnAccompaniedMinorsFlag = RecordsetHelper.ToInt32(rs, "unaccompanied_minor_flag");
                        seatMap.WindowFlag = RecordsetHelper.ToInt32(rs, "window_flag");
                        
                        seatMaps.Add(seatMap);
                        rs.MoveNext();
                    }

                }
                catch
                {
                    throw;
                }
            }

        }
    }
}
