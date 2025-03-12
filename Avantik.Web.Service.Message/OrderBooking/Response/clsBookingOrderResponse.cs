using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ServiceModel;

namespace Avantik.Web.Service.Message.OrderBooking
{
    [MessageContract]
    public class BookingOrderResponse : ResponseBase
    {
        [MessageBodyMember]
        public BookingHeaderResponse BookingHeader { get; set; }
        [MessageBodyMember]
        public IList<FlightSegmentResponse> BookingSegments { get; set; }
        [MessageBodyMember]
        public IList<MappingResponse> Mappings { get; set; }
        [MessageBodyMember]
        public IList<PassengerResponse> Passengers { get; set; }
        [MessageBodyMember]
        public IList<FeeResponse> Fees { get; set; }
        [MessageBodyMember]
        public IList<RemarkResponse> Remarks { get; set; }
        [MessageBodyMember]
        public IList<ServiceResponse> Services { get; set; }
        [MessageBodyMember]
        public IList<PaymentResponse> Payments { get; set; }
        [MessageBodyMember]
        public IList<TaxResponse> Taxes { get; set; }

    }
}
