using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.Serialization;
using System.ServiceModel;

namespace Avantik.Web.Service.Message.Booking
{
    [MessageContract]
    public  class Contact 
    {
        [MessageBodyMember]
        public Address Address { get; set; }
        [MessageBodyMember]
        public Phone Phone { get; set; }
    }
}
