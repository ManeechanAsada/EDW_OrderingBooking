using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.ServiceModel;
using System.Text;

namespace Avantik.Web.Service.Message.Fee
{
    public class BaggageFeeView
    {
        [MessageBodyMember]
        public Guid baggage_fee_option_id { get; set; }
        [MessageBodyMember]
        public Guid passenger_id { get; set; }
        [MessageBodyMember]
        public Guid booking_segment_id { get; set; }
        [MessageBodyMember]
        public Guid fee_id { get; set; }
        [MessageBodyMember]
        public string fee_rcd { get; set; }
        [MessageBodyMember]
        public string fee_category_rcd { get; set; }
        [MessageBodyMember]
        public string currency_rcd { get; set; }
        [MessageBodyMember]
        public string display_name { get; set; }
        [MessageBodyMember]
        public decimal number_of_units { get; set; }
        [MessageBodyMember]
        public decimal fee_amount { get; set; }
        [MessageBodyMember]
        public decimal fee_amount_incl { get; set; }
        [MessageBodyMember]
        public decimal total_amount { get; set; }
        [MessageBodyMember]
        public decimal total_amount_incl { get; set; }
        [MessageBodyMember]
        public decimal vat_percentage { get; set; }      
    }
}
