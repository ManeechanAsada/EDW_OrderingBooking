using Avantik.Web.Service.Contracts;
using Avantik.Web.Service.Message;
using Avantik.Web.Service.Message.SeatMap;
using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.Text;

namespace Avantik.Web.Service.Proxy
{
    public class FlightServiceProxy : ClientBase<IFlightService>, IFlightService
    {
        public GetSeatMapResponse GetSeatMap(GetSeatMapRequest request)
        {
            return base.Channel.GetSeatMap(request);
        }
    }
}
