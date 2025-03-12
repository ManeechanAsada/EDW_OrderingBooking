using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using System.ServiceModel;
using System.Runtime.Serialization;
using Avantik.Web.Service;
using Avantik.Web.Service.Message.Agency;

namespace Avantik.Web.Service.Message
{
      [MessageContract]
    public class AgencySessionProfileResponse : ResponseBase
    {
          [MessageBodyMember]
          public Agent AgentResponse { get; set; }

    }
}
