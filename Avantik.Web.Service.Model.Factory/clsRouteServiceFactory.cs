using Avantik.Web.Service.Helpers;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Configuration;
using System.Linq;
using System.Text;

namespace Avantik.Web.Service.Model.Factory
{
    public static class RouteServiceFactory
    {
        public static Model.Contract.IRouteService CreateInstance()
        {
            bool useCOM = ConfigHelper.ToBoolean("UseCOM");

            if (useCOM == true)
            {
                NameValueCollection Setting = (NameValueCollection)ConfigurationManager.GetSection("ComplusSetting");

                string user = Setting.ToString("ComUser");
                string password = Setting.ToString("ComPassword");
                string server = Setting.ToString("ComServer");
                string domain = Setting.ToString("ComDomain");

                return new Model.COM.RouteService(server, user, password, domain);
            }
            else
            {
                throw new NotImplementedException();

            }

        }
    }
}
