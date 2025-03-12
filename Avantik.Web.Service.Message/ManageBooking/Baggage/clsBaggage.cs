using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.ServiceModel;
using System.Text;

namespace Avantik.Web.Service.Message.ManageBooking
{
    public class Baggage
    {
        [MessageBodyMember]
        public string BookingSegmentID { get; set; }

        [MessageBodyMember]
        public string PassengerID { get; set; }

    }
}
