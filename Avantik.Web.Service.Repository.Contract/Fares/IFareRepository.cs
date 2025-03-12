using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Avantik.Web.Service.Entity.Fares;

namespace Avantik.Web.Service.Repository.Contract.Fares
{
    public interface IFareRepository
    {
        FareLogic GetFareLogicBookingClass(string originRcd,
                                           string destinationRcd,
                                           DateTime dateOutbound,
                                           DateTime dateReturn,
                                           DateTime bookingDate);
    }
}
