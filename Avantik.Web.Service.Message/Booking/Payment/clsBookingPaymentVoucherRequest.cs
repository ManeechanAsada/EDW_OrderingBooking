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
    public class BookingPaymentVoucherRequest
    {
        [MessageBodyMember]
        public BookingMessage Booking { get; set; }
        //[MessageBodyMember]
        //public IList<Voucher> Vouchers { get; set; }
        //[MessageBodyMember]
        //public string RecordLocator { get; set; }
        //[MessageBodyMember]
        //public string ClientProfileId { get; set; }
        //[MessageBodyMember]
        //public bool IncludeOpenVoucher { get; set; }
        //[MessageBodyMember]
        //public bool IncludeExpiredVoucher { get; set; }
        //[MessageBodyMember]
        //public bool IncludeUsedVoucher { get; set; }
        //[MessageBodyMember]
        //public bool IncludeVoidedVoucher { get; set; }
        //[MessageBodyMember]
        //public bool IncludeRefundable { get; set; }
        //[MessageBodyMember]
        //public bool IncludeFareOnly { get; set; }
        //[MessageBodyMember]
        //public bool Write { get; set; }
    }
}
