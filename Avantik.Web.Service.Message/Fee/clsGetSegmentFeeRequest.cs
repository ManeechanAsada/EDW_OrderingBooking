using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using System.ServiceModel;
using System.Runtime.Serialization;
using Avantik.Web.Service;
using Avantik.Web.Service.Message.Booking;

namespace Avantik.Web.Service.Message
{
    [MessageContract]
    public class GetSegmentFeeRequest
    {
        
        [MessageBodyMember]
        public string AgencyCode { set; get; }
        [MessageBodyMember]
        public string CurrencyCode { set; get; }

        [MessageBodyMember]
        public string LanguageCode { set; get; }

        [MessageBodyMember]
        public int NumberOfPassenger { set; get; }

        [MessageBodyMember]
        public int NumberOfInfant { set; get; }
        [MessageBodyMember]
        public bool SpecialService { set; get; }
        [MessageBodyMember]
        public bool bNoVat { set; get; }

        [MessageBodyMember]
        public IList<Message.Booking.PassengerService> services { set; get; }

        [MessageBodyMember]
        public IList<Message.Fee.SegmentService> segmentService { set; get; }

    }
}
