using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.ServiceModel;
using System.Text;

namespace Avantik.Web.Service.Message
{
    [MessageContract]
    public class DestinationsRequest
    {
        [MessageBodyMember]
        public bool B2CFlag { get; set; }
        [MessageBodyMember]
        public bool B2BFlag { get; set; }
        [MessageBodyMember]
        public bool B2EFlag { get; set; }
        [MessageBodyMember]
        public bool B2SFlag { get; set; }
        [MessageBodyMember]
        public bool APIFlag { get; set; }
        [MessageBodyMember]
        public string Language { get; set; }
    }
}
