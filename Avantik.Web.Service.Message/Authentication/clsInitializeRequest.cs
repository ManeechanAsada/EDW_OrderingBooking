using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using System.ServiceModel;
using System.Runtime.Serialization;
using Avantik.Web.Service;

namespace Avantik.Web.Service.Message
{
    [MessageContract]
    public class InitializeRequest
    {
        [MessageBodyMember]
        public string AgencyCode { set; get; }
        [MessageBodyMember]
        public string AgentLogon { set; get; }
        [MessageBodyMember]
        public string AgentPassword { set; get; }

    }
}
