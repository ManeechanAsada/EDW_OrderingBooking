using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ServiceModel;

namespace Avantik.Web.Service.Message.OrderBooking
{
    [MessageContract]
    public class BookingHeader
    {
        [MessageBodyMember]
        public Telephone telephone { get; set; }
        [MessageBodyMember]
        public Address address { get; set; }
        [MessageBodyMember]
        public string agency_code { get; set; }
        [MessageBodyMember]
        public string currency_rcd { get; set; }
        [MessageBodyMember]
        public string booking_number { get; set; }
        [MessageBodyMember]
        public string language_rcd { get; set; }
        [MessageBodyMember]
        public string contact_name { get; set; }
        [MessageBodyMember]
        public string contact_email { get; set; }
        [MessageBodyMember]
        public string received_from { get; set; }
        [MessageBodyMember]
        public string phone_search { get; set; }
        [MessageBodyMember]
        public string comment { get; set; }
        [MessageBodyMember]
        public string title_rcd { get; set; }
        [MessageBodyMember]
        public string lastname { get; set; }
        [MessageBodyMember]
        public string firstname { get; set; }
        [MessageBodyMember]
        public string middlename { get; set; }
        [MessageBodyMember]
        public string country_rcd { get; set; }
        [MessageBodyMember]
        public byte group_booking_flag { get; set; }
    }
}
