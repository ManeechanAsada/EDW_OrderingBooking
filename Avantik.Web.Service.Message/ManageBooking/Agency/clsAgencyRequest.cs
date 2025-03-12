using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.Text;

namespace Avantik.Web.Service.Message
{
    [MessageContract]
    public class AgencyRequest
    {
        [MessageHeader]
        public string Token { get; set; }

    }
}
