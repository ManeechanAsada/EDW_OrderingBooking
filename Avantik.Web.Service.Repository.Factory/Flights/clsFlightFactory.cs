using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using Avantik.Web.Service.Helpers;
using Avantik.Web.Service.Infrastructure;
using Avantik.Web.Service.Repository.Flight;
using Avantik.Web.Service.Repository.Contract;

namespace Avantik.Web.Service.Repository.Factory
{
    public static class FlightFactory
    {
        public static Contract.Flight.IFlightRepository CreateInstance()
        {
            string strSQLConnectionString = ConfigHelper.ToString("SQLConnectionString");
            return new FlightRepository(strSQLConnectionString);
        }
    }
}
