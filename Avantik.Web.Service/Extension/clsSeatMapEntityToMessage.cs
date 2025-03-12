using Avantik.Web.Service.Message;
using Avantik.Web.Service.Message.SeatMap;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Avantik.Web.Service.Extension
{
    public static class SeatMapEntityToMessage
    {
        public static SeatMap ToSeatMapMessage(this Avantik.Web.Service.Entity.Flight.SeatMap map)
        {
            SeatMap seatMap = null;
            if (map != null)
            {
                seatMap = new SeatMap();
                seatMap.AircraftConfigurationCode = map.AircraftConfigurationCode;
                seatMap.AisleFlag = map.AisleFlag;
                seatMap.BassinetFlag = map.BassinetFlag;
                seatMap.BlockB2bFlag = map.BlockB2bFlag;
                seatMap.BlockB2cFlag = map.BlockB2cFlag;
                seatMap.BlockedFlag = map.BlockedFlag;
                seatMap.BoardingClassRcd = map.BoardingClassRcd;
                seatMap.DestinationRcd = map.DestinationRcd;
                seatMap.EmergencyExitFlag = map.EmergencyExitFlag;
                seatMap.FlightCheckInStatusRcd = map.FlightCheckInStatusRcd;
                seatMap.FlightId = map.FlightId;
                seatMap.FreeSeatingFlag = map.FreeSeatingFlag;
                seatMap.HanddicappedFlag = map.HanddicappedFlag;
                seatMap.InfantFlag = map.InfantFlag;
                seatMap.LayoutColumn = map.LayoutColumn;
                seatMap.LayoutRow = map.LayoutRow;
                seatMap.LocationTypeRcd = map.LocationTypeRcd;
                seatMap.LowComfortFlag = map.LowComfortFlag;
                seatMap.NoChildFlag = map.NoChildFlag;
                seatMap.NoInfantFlag = map.NoInfantFlag;
                seatMap.NumberOfBays = map.NumberOfBays;
                seatMap.NumberOfColumns = map.NumberOfColumns;
                seatMap.NumberOfRows = map.NumberOfRows;
                seatMap.OriginRcd = map.OriginRcd;
                seatMap.PassengerCount = map.PassengerCount;
                seatMap.SeatColumn = map.SeatColumn;
                seatMap.SeatRow = map.SeatRow;
                seatMap.FeeRcd = map.FeeRcd;
                seatMap.StretcherFlag = map.StretcherFlag;
                seatMap.UnAccompaniedMinorsFlag = map.UnAccompaniedMinorsFlag;
                seatMap.WindowFlag = map.WindowFlag;
            }
            return seatMap;
        }

        public static IList<Message.SeatMap.SeatMap> ToSeatMapMessage(this IList<Entity.Flight.SeatMap> objEntitySeatMaps)
        {
            List<Message.SeatMap.SeatMap> objMessageSeatMaps = null;
            if (objEntitySeatMaps != null)
            {
                objMessageSeatMaps = new List<Message.SeatMap.SeatMap>();
                for (int i = 0; i < objEntitySeatMaps.Count; i++)
                {
                    objMessageSeatMaps.Add(objEntitySeatMaps[i].ToSeatMapMessage());
                }
            }
            return objMessageSeatMaps;
        }
    }
}
