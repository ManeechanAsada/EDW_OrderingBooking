using System;
using System.Collections.Generic;
using System.Text;
using Avantik.Web.Service.Entity.Route;

namespace Avantik.Web.Service.Model.Contract
{
    public interface IRouteService
    {
        List<Route> GetOrigins(string strLanguage, bool b2cFlag, bool b2bFlag, bool b2EFlag, bool b2SFlag, bool APIFlag);
        List<Route> GetDestinations(string strLanguage, bool b2cFlag, bool b2bFlag, bool b2EFlag, bool b2SFlag, bool APIFlag);

    }
}
