using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.Serialization;
using System.ServiceModel;

namespace Avantik.Web.Service.Message.ManageBooking
{
    public class Fee
    {
        [MessageBodyMember]
        public Guid BookingId { get; set; }
        [MessageBodyMember]
        public Guid BookingSegmentId { get; set; }
        [MessageBodyMember]
        public string Comment { get; set; }
        [MessageBodyMember]
        public string CurrencyRcd { get; set; }
        [MessageBodyMember]
        public string DestinationRcd { get; set; }
        [MessageBodyMember]
        public decimal FeeAmount { get; set; }
        [MessageBodyMember]
        public decimal FeeAmountIncl { get; set; }
        [MessageBodyMember]
        public string FeeRcd { get; set; }
        [MessageBodyMember]
        public string OriginRcd { get; set; }
        [MessageBodyMember]
        public Guid PassengerId { get; set; }
    }
}
