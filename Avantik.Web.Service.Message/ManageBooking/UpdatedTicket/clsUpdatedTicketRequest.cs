using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.Text;

namespace Avantik.Web.Service.Message.ManageBooking
{
    [MessageContract]
    public class UpdatedTicketRequest
    {
        [MessageHeader]
        public string Token { get; set; }

        [MessageBodyMember]
        public string BookingId { set; get; }
        [MessageBodyMember]
        public IList<UpdatedTicket> UpdatedTicket { get; set; }
    }
}
