using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.Serialization;
using System.ServiceModel;

namespace Avantik.Web.Service.Message.Booking
{
    [MessageContract]
    public class BookingMessage
    {
        [MessageBodyMember]
        public BookingHeader Header { get; set; }
        [MessageBodyMember]
        public Contact Contact { get; set; }
        [MessageBodyMember]
        public IList<FlightSegment> FlightSegments { get; set; }
        [MessageBodyMember]
        public IList<Passenger> Passengers { get; set; }
        [MessageBodyMember]
        public IList<Mapping> Mappings { get; set; }
        [MessageBodyMember]
        public IList<PassengerService> Services { get; set; }
        [MessageBodyMember]
        public IList<Fee> Fees { get; set; }
        [MessageBodyMember]
        public IList<Tax> Taxs { get; set; }
        [MessageBodyMember]
        public IList<Payment> Payments { get; set; }
        [MessageBodyMember]
        public IList<Remark> Remarks { get; set; }
        [MessageBodyMember]
        public IList<Quote> Quotes { get; set; }
    }
}
