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
    public class GetSpecialServiceResponse : ResponseBase
    {
          [MessageBodyMember]
          public IList<Message.SpecialService> SpecialServices { get; set; }

    }
}
