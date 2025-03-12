using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.Serialization;
using System.ServiceModel;

namespace Avantik.Web.Service.Message.Booking
{
    public class Address 
    {
        #region Property
        [MessageBodyMember]
        public string AddressLine1  { get; set; }
        [MessageBodyMember]
        public string AddressLine2 { get; set; }
        [MessageBodyMember]
        public string Street { get; set; }
        [MessageBodyMember]
        public string PoBox { get; set; }
        [MessageBodyMember]
        public string City { get; set; }
        [MessageBodyMember]
        public string State { get; set; }
        [MessageBodyMember]
        public string District { get; set; }
        [MessageBodyMember]
        public string Province { get; set; }
        [MessageBodyMember]
        public string ZipCode { get; set; }
        [MessageBodyMember]
        public string CountryRccd { get; set; }

        #endregion
    }
}
