using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using System.ServiceModel;
using System.Runtime.Serialization;
using Avantik.Web.Service;

namespace Avantik.Web.Service.Message.Client
{
    [MessageContract]
    public class CreateClientProfileResponse : ResponseBase
    {
        [MessageBodyMember]
        public Guid ClientProfileId { get; set; }
        
        [MessageBodyMember]
        public string ClientNumber { get; set; }

    }
}
