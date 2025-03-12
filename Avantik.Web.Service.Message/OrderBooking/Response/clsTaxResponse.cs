using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ServiceModel;
namespace Avantik.Web.Service.Message.OrderBooking
{
    [MessageContract]
    public class TaxResponse : Message.OrderBooking.Tax
    {
        [MessageBodyMember]
        public string error_code { get; set; }
        [MessageBodyMember]
        public string error_message { get; set; }
        [MessageBodyMember]
        public string display_name { get; set; }
    }
}
