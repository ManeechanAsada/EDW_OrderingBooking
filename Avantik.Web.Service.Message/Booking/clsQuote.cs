using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.Serialization;
using System.ServiceModel;

namespace Avantik.Web.Service.Message.Booking
{
    public class Quote 
    {
        [MessageBodyMember]
        public Guid BookingSegmentId { get; set; }
        [MessageBodyMember]
        public string PassengerTypeRcd { get; set; }
        [MessageBodyMember]
        public string CurrencyRcd { get; set; }
        [MessageBodyMember]
        public string ChargeType { get; set; }
        [MessageBodyMember]
        public string ChargeName { get; set; }
        [MessageBodyMember]
        public decimal ChargeAmount { get; set; }
        [MessageBodyMember]
        public decimal TotalAmount { get; set; }
        [MessageBodyMember]
        public decimal TaxAmount { get; set; }
        [MessageBodyMember]
        public int PassengerCount { get; set; }
        [MessageBodyMember]
        public int SortSequence { get; set; }
        [MessageBodyMember]
        public Guid CreateBy { get; set; }
        [MessageBodyMember]
        public DateTime CreateDateTime { get; set; }
        [MessageBodyMember]
        public Guid UpdateBy { get; set; }
        [MessageBodyMember]
        public DateTime UpdateDateTime { get; set; }
        [MessageBodyMember]
        public decimal RedemptionPoints { get; set; }
        [MessageBodyMember]
        public decimal ChargeAmountIncl { get; set; }

    }
}
