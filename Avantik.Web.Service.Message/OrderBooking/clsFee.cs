using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ServiceModel;

namespace Avantik.Web.Service.Message.OrderBooking
{
    [MessageContract]
    public class Fee
    {
        [MessageBodyMember]
        public Guid passenger_id { get; set; }
        [MessageBodyMember]
        public Guid booking_segment_id { get; set; }
        [MessageBodyMember]
        public Guid passenger_segment_service_id { get; set; }
        [MessageBodyMember]
        public Guid fee_id { get; set; }

        [MessageBodyMember]
        public string fee_rcd { get; set; }
        [MessageBodyMember]
        public string fee_category_rcd { get; set; }
        [MessageBodyMember]
        public string vendor_rcd { get; set; }
        [MessageBodyMember]
        public string od_origin_rcd { get; set; }
        [MessageBodyMember]
        public string od_destination_rcd { get; set; }
        [MessageBodyMember]
        public string currency_rcd { get; set; }
        [MessageBodyMember]
        public string charge_currency_rcd { get; set; }
        [MessageBodyMember]
        public string unit_rcd { get; set; }
        [MessageBodyMember]
        public string mpd_number { get; set; }
        [MessageBodyMember]
        public string comment { get; set; }
        [MessageBodyMember]
        public string external_reference { get; set; }
        [MessageBodyMember]
        public decimal fee_amount { get; set; }
        [MessageBodyMember]
        public decimal fee_amount_incl { get; set; }
        [MessageBodyMember]
        public decimal vat_percentage { get; set; }
        [MessageBodyMember]
        public decimal charge_amount { get; set; }
        [MessageBodyMember]
        public decimal charge_amount_incl { get; set; }
        [MessageBodyMember]
        public decimal weight_lbs { get; set; }
        [MessageBodyMember]
        public decimal weight_kgs { get; set; }
        [MessageBodyMember]
        public decimal number_of_units { get; set; }
        
        
    }
}
