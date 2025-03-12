using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using System.Collections.Specialized;
using Avantik.Web.Service.Helpers;
using Avantik.Web.Service.Infrastructure;
using Avantik.Web.Service.Repository.Contract;

namespace Avantik.Web.Service.Repository.Factory
{
    public static class AvailabilityFactory
    {
        public static Contract.Flight.IAvailabilityRepository CreateInstance(AvailabilityTypes type)
        {
            if (type == AvailabilityTypes.OWN)
            {
                string strSQLConnectionString = ConfigHelper.ToString("SQLConnectionString");
                return new Flight.AvailabilityRepository(strSQLConnectionString);
            }
            else
            {
                throw new NotImplementedException();
            }
        }
    }
}
