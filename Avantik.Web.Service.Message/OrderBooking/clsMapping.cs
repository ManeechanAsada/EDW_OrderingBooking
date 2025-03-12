using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ServiceModel;

namespace Avantik.Web.Service.Message.OrderBooking
{
    [MessageContract]
    public class Mapping
    {
        [MessageBodyMember]
        public Guid passenger_id { get; set; }
        [MessageBodyMember]
        public Guid booking_segment_id { get; set; }
        [MessageBodyMember]
        public DateTime not_valid_before_date { get; set; }
        [MessageBodyMember]
        public DateTime not_valid_after_date { get; set; }
        [MessageBodyMember]
        public Int16 refund_with_charge_hours { get; set; }
        [MessageBodyMember]
        public Int16 refund_not_possible_hours { get; set; }
        [MessageBodyMember]
        public short piece_allowance { get; set; }
        [MessageBodyMember]
        public string fare_code { get; set; }
        [MessageBodyMember]
        public string seat_number { get; set; }
        [MessageBodyMember]
        public string currency_rcd { get; set; }
        [MessageBodyMember]
        public string endorsement_text { get; set; }
        [MessageBodyMember]
        public string restriction_text { get; set; }
        [MessageBodyMember]
        public decimal fare_amount { get; set; }
        [MessageBodyMember]
        public decimal fare_amount_incl { get; set; }
        [MessageBodyMember]
        public decimal fare_vat { get; set; }
        [MessageBodyMember]
        public decimal net_total { get; set; }
        //Optinal
        [MessageBodyMember]
        public decimal commission_amount { get; set; }
        //Optinal
        [MessageBodyMember]
        public decimal commission_amount_incl { get; set; }
        //Optinal
        [MessageBodyMember]
        public decimal commission_percentage { get; set; }
        //Optinal
        [MessageBodyMember]
        public decimal public_fare_amount { get; set; }
        //Optinal
        [MessageBodyMember]
        public decimal public_fare_amount_incl { get; set; }
        [MessageBodyMember]
        public decimal refund_charge { get; set; }
        [MessageBodyMember]
        public decimal baggage_weight { get; set; }
        [MessageBodyMember]
        public double redemption_points { get; set; }
        [MessageBodyMember]
        public byte e_ticket_flag { get; set; }
        [MessageBodyMember]
        public byte refundable_flag { get; set; }
        [MessageBodyMember]
        public byte transferable_fare_flag { get; set; }
        [MessageBodyMember]
        public byte through_fare_flag { get; set; }
        [MessageBodyMember]
        public byte it_fare_flag { get; set; }
        [MessageBodyMember]
        public byte duty_travel_flag { get; set; }
        [MessageBodyMember]
        public byte standby_flag { get; set; }
        [MessageBodyMember]
        public byte exclude_pricing_flag { get; set; }


    }
}
