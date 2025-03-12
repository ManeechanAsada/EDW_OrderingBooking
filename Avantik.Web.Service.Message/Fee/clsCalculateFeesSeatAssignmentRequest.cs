using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using System.ServiceModel;
using System.Runtime.Serialization;
using Avantik.Web.Service;
using Avantik.Web.Service.Message.Booking;

namespace Avantik.Web.Service.Message
{
    [MessageContract]
    public class CalculateFeesSeatAssignmentRequest
    {
        [MessageBodyMember]
        public BookingRequest Booking { set; get; }
        [MessageBodyMember]
        public string AgencyCode { set; get; }
        [MessageBodyMember]
        public string Currency { set; get; }
        [MessageBodyMember]
        public string Language { set; get; }
        [MessageBodyMember]
        public bool bNovat { set; get; }

    }
}
