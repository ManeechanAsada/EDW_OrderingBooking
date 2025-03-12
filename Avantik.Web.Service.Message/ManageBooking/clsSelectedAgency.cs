using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.ServiceModel;
using System.Text;

namespace Avantik.Web.Service.Message.ManageBooking
{
    public class SelectedAgency
    {
        [MessageBodyMember]
        public string UserId { get; set; }
        [MessageBodyMember]
        public string AgencyCode { get; set; }
        [MessageBodyMember]
        public string CurrencyRcd { get; set; }
        [MessageBodyMember]
        public string LanguageRcd { get; set; }
    }
}
