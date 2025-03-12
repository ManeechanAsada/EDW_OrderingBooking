using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.Text;

namespace Avantik.Web.Service.Message
{
    [MessageContract]
    public class NameChangeRequest
    {
        [MessageHeader]
        public string Token { get; set; }

        [MessageBodyMember]
        public string BookingId { set; get; }
        [MessageBodyMember]
        public IList<Message.ManageBooking.NameChange> NameChange { get; set; }
    }
}
