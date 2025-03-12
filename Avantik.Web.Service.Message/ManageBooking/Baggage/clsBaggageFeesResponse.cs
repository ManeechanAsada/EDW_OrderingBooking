using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.Text;
using Avantik.Web.Service.Message.Booking;

namespace Avantik.Web.Service.Message.ManageBooking
{
    [MessageContract]
    public class BaggageFeesResponse : ResponseBase
    {
        [MessageBodyMember]
        public BookingResponse BookingResponse { get; set; }
    }
}
