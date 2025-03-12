using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Avantik.Web.Service.Helpers;
using Avantik.Web.Service.Infrastructure;
using Avantik.Web.Service.Repository.Fares;
using Avantik.Web.Service.Repository.Contract;

namespace Avantik.Web.Service.Repository.Factory.Fares
{
    public static class FareFactory
    {
        public static Contract.Fares.IFareRepository CreateInstance()
        {
            string strSQLConnectionString = ConfigHelper.ToString("SQLConnectionString");
            return new FareRepository(strSQLConnectionString);
        }
    }
}
