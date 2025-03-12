using Avantik.Web.Service.Message.Agency;
using Avantik.Web.Service.Message.Fee;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Avantik.Web.Service.Extension
{
    public static class FeeMessageToEntity
    {
        public static IList<Entity.SegmentService> ToFeeEntity(this IList<Message.Fee.SegmentService> objEntitytFeeList)
        {
            List<Entity.SegmentService> objMessageFeeList = null;
            if (objEntitytFeeList != null)
            {
                objMessageFeeList = new List<Entity.SegmentService>();
                for (int i = 0; i < objEntitytFeeList.Count; i++)
                {
                    objMessageFeeList.Add(objEntitytFeeList[i].ToFeeEntity());
                }
            }
            return objMessageFeeList;
        }

        public static Entity.SegmentService ToFeeEntity(this Message.Fee.SegmentService sf)
        {
            Entity.SegmentService segmentFee = null;

            if(sf != null)
            {
                segmentFee = new Entity.SegmentService();
                segmentFee.FlightConnectionId = sf.FlightConnectionId;
                segmentFee.SpecialServiceRcd = sf.SpecialServiceRcd;
                segmentFee.OriginRcd = sf.OriginRcd;
                segmentFee.DestinationRcd = sf.DestinationRcd;
                segmentFee.OdOriginRcd = sf.OdOriginRcd;
                segmentFee.OdDestinationRcd = sf.OdDestinationRcd;
                segmentFee.BookingClassRcd = sf.BookingClassRcd;
                segmentFee.FareCode = sf.FareCode;
                segmentFee.AirlineRcd = sf.AirlineRcd;
                segmentFee.FlightNumber = sf.FlightNumber;
                segmentFee.DepartureDate = sf.DepartureDate;

            }
            return segmentFee;
        }

    }
}
