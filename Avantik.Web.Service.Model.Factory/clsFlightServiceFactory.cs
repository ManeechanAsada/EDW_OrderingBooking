using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using Avantik.Web.Service;
using Avantik.Web.Service.Helpers;
using Avantik.Web.Service.Infrastructure;
using Avantik.Web.Service.Model.COM;
using System.Collections.Specialized;
using System.Configuration;

namespace Avantik.Web.Service.Model
{
    public static class FlightServiceFactory
    {
        public static Model.Contract.IAvailabilityService CreateAvailabilityInstance(AvailabilityTypes availabilityType)
        {
            bool useCOM = ConfigHelper.ToBoolean("UseCOMAvailability");

            if (useCOM == false)
            {
                //Create Repository instance.
                Repository.Contract.Flight.IAvailabilityRepository availabilityRepository = Repository.Factory.AvailabilityFactory.CreateInstance(availabilityType);
                Repository.Contract.Flight.IFlightRepository flightRepository = Repository.Factory.FlightFactory.CreateInstance();

                //Create Service model instance.
                return new Model.AvailabilityService(availabilityType, availabilityRepository, flightRepository);
            }
            else
            {
                if (availabilityType == AvailabilityTypes.OWN)
                {
                    //Create Service model instance.
                    NameValueCollection Setting = (NameValueCollection)ConfigurationManager.GetSection("ComplusSetting");

                    string user = Setting.ToString("ComUser");
                    string password = Setting.ToString("ComPassword");
                    string server = Setting.ToString("ComServer");
                    string domain = Setting.ToString("ComDomain");

                    return new Model.COM.Flight.AvailabilityService(server, user, password, domain);
                }
                else
                {
                    return null;
                }
                
            }
            
        }

        public static Model.Contract.IFlightService CreateInstance()
        {
            bool useCOM = ConfigHelper.ToBoolean("UseCOM");

            if (useCOM == true)
            {
                NameValueCollection Setting = (NameValueCollection)ConfigurationManager.GetSection("ComplusSetting");

                string user = Setting.ToString("ComUser");
                string password = Setting.ToString("ComPassword");
                string server = Setting.ToString("ComServer");
                string domain = Setting.ToString("ComDomain");

                return new Model.COM.FlightService(server, user, password, domain);
            }
            else
            {
                throw new NotImplementedException();

            }
        }
    }
}
