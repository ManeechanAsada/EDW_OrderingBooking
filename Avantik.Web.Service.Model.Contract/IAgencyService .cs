using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Avantik.Web.Service.Entity.Flight;
using Avantik.Web.Service.Infrastructure;
using Avantik.Web.Service.Entity.Booking;
using Avantik.Web.Service.Entity.Agency;

namespace Avantik.Web.Service.Model.Contract
{
    public interface IAgencyService
    {
        Agent GetAgencySessionProfile(string strAgencyCode,
                                       string strUserId);
        Agent TravelAgentLogon(string AgencyCode,
                                string AgencyLogon,
                                string AgencyPassword);

    }
}
