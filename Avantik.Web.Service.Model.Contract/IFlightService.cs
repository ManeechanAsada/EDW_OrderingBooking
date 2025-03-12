using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Avantik.Web.Service.Infrastructure;
using Avantik.Web.Service.Entity.Flight;

namespace Avantik.Web.Service.Model.Contract
{
    public interface IFlightService
    {
        List<SeatMap> GetSeatMap(Entity.Booking.Flight flight);
    }
}
