using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ServiceModel;

namespace Avantik.Web.Service.Message.OrderBooking
{
    [MessageContract]
    public class Tax
    {
        [MessageBodyMember]
        public Guid passenger_id { get; set; }
        [MessageBodyMember]
        public Guid booking_segment_id { get; set; }
        [MessageBodyMember]
        public string tax_rcd { get; set; }
        [MessageBodyMember]
        public string tax_currency_rcd { get; set; }
        [MessageBodyMember]
        public string sales_currency_rcd { get; set; }
        [MessageBodyMember]
        public decimal sales_amount { get; set; }
        [MessageBodyMember]
        public decimal sales_amount_incl { get; set; }
        [MessageBodyMember]
        public decimal tax_amount { get; set; }
        [MessageBodyMember]
        public decimal tax_amount_incl { get; set; }
    }
}
