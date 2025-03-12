using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using System.ServiceModel;
using System.Runtime.Serialization;
using Avantik.Web.Service;

namespace Avantik.Web.Service.Message
{
    [MessageContract]
    public class GetFeeDefinitionRequest
    {
        [MessageHeader]
        public string Token { get; set; }

        [MessageBodyMember]
        public string FeeRcd { set; get; }
        [MessageBodyMember]
        public string BookingClass { set; get; }
        [MessageBodyMember]
        public string FareCode { set; get; }
        [MessageBodyMember]
        public string Origin { set; get; }
        [MessageBodyMember]
        public string Destination { set; get; }
        [MessageBodyMember]
        public string FlightNumber { set; get; }
        [MessageBodyMember]
        public string Currency { set; get; }
    }
}
