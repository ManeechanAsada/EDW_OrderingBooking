using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using System.ServiceModel;
using System.Runtime.Serialization;
using Avantik.Web.Service;

namespace Avantik.Web.Service.Message.Booking
{
    [MessageContract]
    public class BookingFlightRequest
    {
        [MessageHeader]
        public string Token { get; set; }

        [MessageBodyMember]
        public string AgencyCode { get; set; }

        [MessageBodyMember]
        public string Currency { get; set; }

        [MessageBodyMember]
        public IList<Flight> Flight { get; set; }

        [MessageBodyMember]
        public string BookingId { get; set; }

        [MessageBodyMember]
        public short Adults { get; set; }

        [MessageBodyMember]
        public short Children { get; set; }

        [MessageBodyMember]
        public short Infants { get; set; }

        [MessageBodyMember]
        public short Others { get; set; }

        [MessageBodyMember]
        public string StrOthers { get; set; }

        [MessageBodyMember]
        public string UserId { get; set; }

        [MessageBodyMember]
        public string StrIpAddress { get; set; }

        [MessageBodyMember]
        public string StrLanguageCode { get; set; }

        [MessageBodyMember]
        public bool BNoVat { get; set; }

    }
}
