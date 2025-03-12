using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.ServiceModel;

namespace Avantik.Web.Service.Message.Booking
{
    [MessageContract]
    public class QuoteSummaryResponse : ResponseBase
    {
        [MessageBodyMember]
        public BookingResponse BookingResponse { get; set; }
    }
}
