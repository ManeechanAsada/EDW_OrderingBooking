using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.ServiceModel;
namespace Avantik.Web.Service.Message.Booking
{
    [MessageContract]
    public class QuoteSummaryRequest
    {
        [MessageBodyMember]
        public IList<FlightSegment> FlightSegments { get; set; }
        [MessageBodyMember]
        public IList<Passenger> Passengers { get; set; }
        [MessageBodyMember]
        public string AgencyCode { get; set; }
        [MessageBodyMember]
        public string Language { get; set; }
        [MessageBodyMember]
        public string CurrencyCode { get; set; }
        [MessageBodyMember]
        public bool bNovat { get; set; }
    }
}
