using Avantik.Web.Service.Helpers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Avantik.Web.Service.Model.COM;
using System.Collections.Specialized;
using System.Configuration;

namespace Avantik.Web.Service.Model.Booking
{
    public static class PaymentServiceFactory 
    {
        public static Model.Contract.IPaymentService CreateInstance()
        {
            bool useCOM = ConfigHelper.ToBoolean("UseCOM");

            if (useCOM == true)
            {
                NameValueCollection Setting = (NameValueCollection)ConfigurationManager.GetSection("ComplusSetting");

                string user = Setting.ToString("ComUser");
                string password = Setting.ToString("ComPassword");
                string server = Setting.ToString("ComServer");
                string domain = Setting.ToString("ComDomain");

                return new Model.COM.PaymentService(server, user, password, domain);
            }
            else
            {
                throw new NotImplementedException();

            }


        }

    }
}
