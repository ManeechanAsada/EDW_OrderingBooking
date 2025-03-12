using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.ServiceModel;
using System.Text;

namespace Avantik.Web.Service.Message.ManageBooking
{
    public class UpdatedTicket
    {
        [MessageBodyMember]
        public string BookingSegmentID { get; set; }
        [MessageBodyMember]
        public string PassengerID { get; set; }
        [MessageBodyMember]
        public decimal WeightAllowance { get; set; }
        [MessageBodyMember]
        public short PieceAllowance { get; set; }
        [MessageBodyMember]
        public string Endorsegment { get; set; }
        [MessageBodyMember]
        public string Restriction { get; set; }

    }
}
