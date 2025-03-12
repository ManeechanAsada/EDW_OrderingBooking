using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ServiceModel;

namespace Avantik.Web.Service.Message.OrderBooking
{
    [MessageContract]
    public class Telephone
    {
        [MessageBodyMember]
        public string phone_mobile { get; set; }
        [MessageBodyMember]
        public string phone_home { get; set; }
        [MessageBodyMember]
        public string phone_business { get; set; }
        [MessageBodyMember]
        public string phone_fax { get; set; }
    }
}
