using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Avantik.Web.Service.Entity.Booking

{
    public class Quote
    {
        public Guid BookingSegmentId { get; set; }
        public string PassengerTypeRcd { get; set; }
        public string CurrencyRcd { get; set; }
        public string ChargeType { get; set; }
        public string ChargeName { get; set; }
        public decimal ChargeAmount { get; set; }
        public decimal TotalAmount { get; set; }
        public decimal TaxAmount { get; set; }
        public int PassengerCount { get; set; }
        public int SortSequence { get; set; }
        public Guid CreateBy { get; set; }
        public DateTime CreateDateTime { get; set; }
        public Guid UpdateBy { get; set; }
        public DateTime UpdateDateTime { get; set; }
        public decimal RedemptionPoints { get; set; }
        public decimal ChargeAmountIncl { get; set; }

    }
}
