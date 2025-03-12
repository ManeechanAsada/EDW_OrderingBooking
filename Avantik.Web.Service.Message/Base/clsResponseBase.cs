using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.Serialization;
using System.ServiceModel;

namespace Avantik.Web.Service.Message
{
    [MessageContract]
    public abstract class ResponseBase
    {
        [MessageBodyMember]
        public bool Success { get; set; }
        [MessageBodyMember]
        public string Message { get; set; }
        [MessageBodyMember]
        public string Code { get; set; }

    }
}
