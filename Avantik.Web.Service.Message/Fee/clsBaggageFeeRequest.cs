using Avantik.Web.Service.Message.Booking;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.ServiceModel;
using System.Text;

namespace Avantik.Web.Service.Message.Fee
{
    [MessageContract]
    public class BaggageFeeRequest
    {
        [MessageBodyMember]
        public Guid BookingSegmentId { get; set; }
        [MessageBodyMember]
        public Guid PassengerId { get; set; }
        [MessageBodyMember]
        public string AgencyCode { get; set; }
        [MessageBodyMember]
        public string LanguageCode { get; set; }
        [MessageBodyMember]
        public IList<Mapping> Mappings { get; set; }
        [MessageBodyMember]
        public int MaxUnit { get; set; }
        [MessageBodyMember]
        public IList<Message.Booking.Fee> BookingFees { get; set; }

        [MessageBodyMember]
        public bool bNovat { get; set; }
    }
}
