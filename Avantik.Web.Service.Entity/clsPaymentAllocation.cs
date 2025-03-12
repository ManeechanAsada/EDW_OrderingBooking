using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Avantik.Web.Service.Entity
{
    public class PaymentAllocation
    {
        public Guid booking_payment_id { get; set; }
        public Guid passenger_id { get; set; }
        public Guid booking_segment_id { get; set; }
        public Guid booking_fee_id { get; set; }
        public Guid voucher_id { get; set; }
        public Guid passenger_segment_service_id { get; set; }
        public Guid fee_id { get; set; }
        public Guid user_id { get; set; }

        public string currency_rcd { get; set; }
        public string charge_currency_rcd { get; set; }
        public string od_origin_rcd { get; set; }
        public string od_destination_rcd { get; set; }
        public string fee_category_rcd { get; set; }
        public string vendor_rcd { get; set; }
        public string units { get; set; }
        public string external_reference { get; set; }

        public decimal sales_amount { get; set; }
        public decimal payment_amount { get; set; }
        public decimal account_amount { get; set; }
        public decimal charge_amount { get; set; }
        public decimal charge_amount_incl { get; set; }
        public decimal weight_lbs { get; set; }
        public decimal weight_kgs { get; set; }
    }
}
