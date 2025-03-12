using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.ServiceModel;
using System.Text;
namespace Avantik.Web.Service.Message.ManageBooking
{
    public class PaymentFee
    {
        [MessageBodyMember]
        public string FormOfPaymentRcd { get; set; }
        [MessageBodyMember]
        public string FormOfPaymentSubtypeRcd { get; set; }
        [MessageBodyMember]
        public string CurrencyRcd { get; set; }
        [MessageBodyMember]
        public string AgencyCode { get; set; }
        [MessageBodyMember]
        public string Origin { get; set; }
        [MessageBodyMember]
        public string Destination { get; set; }

    }
}
