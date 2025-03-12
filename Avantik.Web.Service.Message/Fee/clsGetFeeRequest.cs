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
    public class GetFeeRequest
    {
        [MessageBodyMember]
        public string StrFeeRcd { set; get; }
        [MessageBodyMember]
        public string StrCurrencyCode { set; get; }
        [MessageBodyMember]
        public string StrAgencyCode { set; get; }

        [MessageBodyMember]
        public string StrClass { set; get; }
        [MessageBodyMember]
        public string StrFareBasis { set; get; }
        [MessageBodyMember]
        public string StrOrigin { set; get; }
        [MessageBodyMember]
        public string StrDestination { set; get; }
        [MessageBodyMember]
        public string StrFlightNumber { set; get; }
        [MessageBodyMember]
        public string StrLanguage { set; get; }
        [MessageBodyMember]
        public DateTime DtDate { set; get; }
        [MessageBodyMember]
        public bool bNoVat { set; get; }


    }
}
