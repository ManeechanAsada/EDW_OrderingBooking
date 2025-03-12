using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ServiceModel;

namespace Avantik.Web.Service.Message.OrderBooking
{
    [MessageContract]
    public class Address
    {
        [MessageBodyMember]
        public string address_line1 { get; set; }
        [MessageBodyMember]
        public string address_line2 { get; set; }
        [MessageBodyMember]
        public string street { get; set; }
        [MessageBodyMember]
        public string po_box { get; set; }
        [MessageBodyMember]
        public string city { get; set; }
        [MessageBodyMember]
        public string state { get; set; }
        [MessageBodyMember]
        public string district { get; set; }
        [MessageBodyMember]
        public string province { get; set; }
        [MessageBodyMember]
        public string zip_code { get; set; }
        [MessageBodyMember]
        public string country_rcd { get; set; }
    }
}
