using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.Serialization;
using System.ServiceModel;

namespace Avantik.Web.Service.Message.FormOfPayment
{
    public class FormOfPaymentSubType 
    {
        #region Property
        [MessageBodyMember]
        public Int32 VoucherReference { get; set; }
        [MessageBodyMember]
        public Int16 ExpiryDays { get; set; }
        [MessageBodyMember]
        public byte CvvRequiredFlag { get; set; }
        [MessageBodyMember]
        public byte ValidateDocumentNumberFlag { get; set; }
        [MessageBodyMember]
        public byte DisplayCvvFlag { get; set; }
        [MessageBodyMember]
        public byte MultiplePaymentFlag { get; set; }
        [MessageBodyMember]
        public byte ApprovalCodeRequiredFlag { get; set; }
        [MessageBodyMember]
        public byte DisplayApprovalCodeFlag { get; set; }
        [MessageBodyMember]
        public byte DisplayExpiryDateFlag { get; set; }
        [MessageBodyMember]
        public byte ExpiryDateRequiredFlag { get; set; }
        [MessageBodyMember]
        public byte TravelAgencyPaymentFlag { get; set; }
        [MessageBodyMember]
        public byte AgencyPaymentFlag { get; set; }
        [MessageBodyMember]
        public byte InternetPaymentFlag { get; set; }
        [MessageBodyMember]
        public byte RefundPaymentFlag { get; set; }
        [MessageBodyMember]
        public byte AddressRequiredFlag { get; set; }
        [MessageBodyMember]
        public byte DisplayAddressFlag { get; set; }
        [MessageBodyMember]
        public byte ShowPosIndictorFlag { get; set; }
        [MessageBodyMember]
        public byte RequirePosIndicatorFlag { get; set; }
        [MessageBodyMember]
        public byte DisplayIssueDateFlag { get; set; }
        [MessageBodyMember]
        public byte DisplayIssueNumberFlag { get; set; }
        [MessageBodyMember]
        public string FormOfPaymentSubtypeRcd { get; set; }
        [MessageBodyMember]
        public string DisplayName { get; set; }
        [MessageBodyMember]
        public string FormOfPaymentRcd { get; set; }
        [MessageBodyMember]
        public string CardCode { get; set; }



        #endregion
    }
}
