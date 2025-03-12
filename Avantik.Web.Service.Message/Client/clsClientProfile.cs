using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.Serialization;
using System.ServiceModel;

namespace Avantik.Web.Service.Message.Client
{
    public class ClientProfile 
    {
        [MessageBodyMember]
        public Client Client { get; set; }
        [MessageBodyMember]
        public IList<PassengerProfile> PassengerProfiles { get; set; }
        [MessageBodyMember]
        public IList<Message.Booking.Remark> BookingRemarks { get; set; }

    }
}
