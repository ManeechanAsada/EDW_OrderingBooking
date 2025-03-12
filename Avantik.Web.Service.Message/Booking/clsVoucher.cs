using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.Serialization;
using System.ServiceModel;

namespace Avantik.Web.Service.Message
{
    public class Voucher : VoucherTemplate
    {
        #region Property
        [MessageBodyMember]
        public string AgencyCode { get; set; }
        [MessageBodyMember]
        public Guid CreateBy { get; set; }
        [MessageBodyMember]
        public DateTime CreateDateTime { get; set; }
        [MessageBodyMember]
        public DateTime ExpiryDateTime { get; set; }
        [MessageBodyMember]
        public decimal PaymentTotal { get; set; }
        [MessageBodyMember]
        public byte PercentageFlag { get; set; }
        [MessageBodyMember]
        public string RecipientName { get; set; }
        [MessageBodyMember]
        public byte RefundableFlag { get; set; }
        [MessageBodyMember]
        public byte SingleFlightFlag { get; set; }
        [MessageBodyMember]
        public byte SinglePassengerFlag { get; set; }
        [MessageBodyMember]
        public Guid UpdateBy { get; set; }
        [MessageBodyMember]
        public DateTime UpdateDateTime { get; set; }
        [MessageBodyMember]
        public string VoucherId { get; set; }
        [MessageBodyMember]
        public string VoucherNumber { get; set; }
        [MessageBodyMember]
        public string VoucherPassword { get; set; }
        [MessageBodyMember]
        public string VoucherStatusRcd { get; set; }
        #endregion
    }
}
