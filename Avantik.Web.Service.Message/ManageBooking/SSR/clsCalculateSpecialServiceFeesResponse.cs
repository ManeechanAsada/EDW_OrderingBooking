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
    public class CalculateSpecialServiceFeesResponse : ResponseBase
    {
        [MessageBodyMember]
        public  BookingResponse BookingResponse{ get; set; }
    }
}
