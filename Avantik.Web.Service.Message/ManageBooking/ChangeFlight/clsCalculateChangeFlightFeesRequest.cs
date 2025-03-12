using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using System.ServiceModel;
using System.Runtime.Serialization;
using Avantik.Web.Service;
using Avantik.Web.Service.Message.Booking;
using Avantik.Web.Service.Message.ManageBooking;

namespace Avantik.Web.Service.Message
{
    [MessageContract]
    public class CalculateChangeFlightFeesRequest
    {
        [MessageHeader]
        public string Token { set; get; }
        [MessageBodyMember]
        public string BookingId { set; get; }
        [MessageBodyMember]
        public IList<ChangeFlight> ModifyFlights { get; set; }
    }
}
