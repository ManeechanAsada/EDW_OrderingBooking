using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Avantik.Web.Service.Entity.Booking;

namespace Avantik.Web.Service.Extension
{
    public static class FlightMessageToEntity
    {
        public static Flight ToEntity(this Message.Booking.Flight flightMessage)
        {
            Flight flight = null;
            if (flightMessage != null)
            {
                flight = new Flight();
                flight.BoardingClassRcd = flightMessage.BoardingClassRcd;
                flight.BookingClassRcd = flightMessage.BookingClassRcd;
                flight.DepartureDate = flightMessage.DepartureDate;
                flight.DestinationRcd = flightMessage.DestinationRcd;
                flight.EticketFlag = flightMessage.EticketFlag;
                flight.FairId = flightMessage.FairId;
                flight.FlightConnectionId = flightMessage.FlightConnectionId;
                flight.FlightId = flightMessage.FlightId;
                flight.OdDestinationRcd = flightMessage.OdDestinationRcd;
                flight.OdOriginRcd = flightMessage.OdOriginRcd;
                flight.OriginRcd = flightMessage.OriginRcd;
                flight.TransitPoints = flightMessage.TransitPoints;
                flight.TransitpointsName = flightMessage.TransitPointsName;
                flight.NumberOfUnits = flightMessage.NumberOfUnits;
                flight.AirlineRcd = flightMessage.AirlineRcd;
                flight.FlightNumber = flightMessage.FlightNumber;

            }
            return flight;
        }

        public static IList<Flight> ToListEntity(this  IList<Message.Booking.Flight> segment)
        {
            IList<Flight> segmentList = null;
            if (segment != null)
            {
                segmentList = new List<Flight>();

                for (int i = 0; i < segment.Count; i++)
                {
                    segmentList.Add(segment[i].ToEntity());
                }
            }

            return segmentList;
        }
    }
}
