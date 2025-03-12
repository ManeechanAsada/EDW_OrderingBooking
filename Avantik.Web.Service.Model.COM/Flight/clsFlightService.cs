using Avantik.Web.Service.Model.Contract;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Avantik.Web.Service.COMHelper;
using System.Runtime.InteropServices;
using Avantik.Web.Service.Model.COM.Extension;
using Avantik.Web.Service.Entity;
using Avantik.Web.Service.Entity.Flight;
using Avantik.Web.Service.Entity.Booking;

namespace Avantik.Web.Service.Model.COM
{
    public class FlightService : RunComplus, IFlightService
    {
        string _server = string.Empty;
        public FlightService(string server, string user, string pass, string domain)
            :base(user,pass,domain)
        {
            _server = server;
        }
        public List<SeatMap> GetSeatMap(Entity.Booking.Flight flight)
        {          
            //tikAeroProcess.CheckIn objCheckin = null;
            tikAeroProcess.Flights objFlight = null;
            ADODB.Recordset rs = null;
            //ADODB.Recordset rsAttributeMapping = null;
            List<SeatMap> seatMaps = new List<SeatMap>();
            //string strLanguage = "EN";
            //bool bAttributeMapping = false;
            string strOriginRcd = string.Empty;
            string strDestinationRcd = string.Empty;
            string strFlightId = "00000000-0000-0000-0000-000000000000";
            string strBoardingClassRcd = string.Empty;
            string strBookingClassRcd = string.Empty;

            try
            {
                if (_server.Length > 0)
                {
                    //Type remote = Type.GetTypeFromProgID("tikAeroProcess.CheckIn", _server);
                    //objCheckin = (tikAeroProcess.CheckIn)Activator.CreateInstance(remote);
                    Type remote = Type.GetTypeFromProgID("tikAeroProcess.Flights", _server);
                    objFlight = (tikAeroProcess.Flights)Activator.CreateInstance(remote);
                    remote = null;
                }
                else
                //{ objCheckin = new tikAeroProcess.CheckIn(); }
                { objFlight = new tikAeroProcess.Flights(); }

                strOriginRcd = flight.OriginRcd;
                strDestinationRcd = flight.DestinationRcd;
                if (flight.FlightId != Guid.Empty)
                    strFlightId = flight.FlightId.ToString();
                strBoardingClassRcd = flight.BoardingClassRcd;
                //   strBookingClassRcd = flight.BookingClassRcd;

                //EDWDEV-284
                //rs = objCheckin.GetSeatMap(ref strOriginRcd, 
                //                            ref strDestinationRcd,
                //                            ref strFlightId,
                //                            ref strBoardingClassRcd,
                //                            ref strBookingClassRcd,
                //                            ref strLanguage,
                //                            ref rsAttributeMapping,
                //                            ref bAttributeMapping);
 
                var flightid = flight.FlightId.ToString();
                var origin = flight.OriginRcd;
                var destination = flight.DestinationRcd;
                var boardingClass = flight.BoardingClassRcd;
                var strConfiguration = string.Empty;
                var language = string.Empty;
                ADODB.Recordset rsCompartment = null;
                ADODB.Recordset rsSeatmap = null;
                ADODB.Recordset rsAttribute = null;

                objFlight.GetSeatmapLayout(ref flightid, ref origin, ref destination, ref boardingClass, ref strConfiguration, ref language, ref rsCompartment, ref rsSeatmap, ref rsAttribute);

                if (rsSeatmap != null && rsSeatmap.RecordCount > 0)
                {
                    seatMaps.FillSeatMap(ref rsSeatmap);
                }
            }
            catch (SystemException ex)
            {
                throw ex;
            }
            finally
            {
                if (objFlight != null)
                { Marshal.FinalReleaseComObject(objFlight); }
                objFlight = null;

                RecordsetHelper.ClearRecordset(ref rs);
            }

            return seatMaps;
        }
    }
}
