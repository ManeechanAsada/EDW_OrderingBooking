using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.ServiceModel;
using System.Text;

namespace Avantik.Web.Service.Message.ManageBooking
{
    public class SeatAssign
    {
        [MessageBodyMember]
        public string BookingSegmentID { get; set; }
        [MessageBodyMember]
        public string PassengerID { get; set; }
        [MessageBodyMember]
        public Int32 SeatRow { get; set; }
        [MessageBodyMember]
        public string SeatColumn { get; set; }
       // [MessageBodyMember]
       // public string SeatNumber { get; set; }
        [MessageBodyMember]
        public string SeatFeeRcd { get; set; }

    }
}
