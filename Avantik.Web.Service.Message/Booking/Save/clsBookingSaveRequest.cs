using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using System.ServiceModel;
using System.Runtime.Serialization;
using Avantik.Web.Service;
using Avantik.Web.Service.Message.Agency;

namespace Avantik.Web.Service.Message.Booking
{
    [MessageContract]
    public class BookingSaveRequest
    {
        [MessageHeader]
        public string Token { get; set; }

        [MessageBodyMember]
        public BookingRequest Booking { get; set; }

        [MessageBodyMember]
        public bool CheckSeatAssignment { get; set; }
        
        [MessageBodyMember]
        public bool CheckSessionTimeOut { get; set; }

    }
}
