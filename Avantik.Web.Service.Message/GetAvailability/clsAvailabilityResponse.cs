using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.Serialization;
using System.ServiceModel;

namespace Avantik.Web.Service.Message
{
    [MessageContract]
    public class AvailabilityResponse : ResponseBase
    {
        [MessageBodyMember]
        public IEnumerable<AvailabilityView> AvailabilityOutbound { get; set; }
        [MessageBodyMember]
        public IEnumerable<AvailabilityView> AvailabilityReturn { get; set; }
    }
}
