using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

namespace Avantik.Web.Service.Entity.FormOfPayment
{
    public class FormOfPaymentSubType 
    {
        #region Property

        public Int32 VoucherReference { get; set; }

        public Int16 ExpiryDays { get; set; }

        public byte CvvRequiredFlag { get; set; }
        public byte ValidateDocumentNumberFlag { get; set; }
        public byte DisplayCvvFlag { get; set; }
        public byte MultiplePaymentFlag { get; set; }
        public byte ApprovalCodeRequiredFlag { get; set; }
        public byte DisplayApprovalCodeFlag { get; set; }
        public byte DisplayExpiryDateFlag { get; set; }
        public byte ExpiryDateRequiredFlag { get; set; }
        public byte TravelAgencyPaymentFlag { get; set; }
        public byte AgencyPaymentFlag { get; set; }
        public byte InternetPaymentFlag { get; set; }
        public byte RefundPaymentFlag { get; set; }
        public byte AddressRequiredFlag { get; set; }
        public byte DisplayAddressFlag { get; set; }
        public byte ShowPosIndictorFlag { get; set; }
        public byte RequirePosIndicatorFlag { get; set; }
        public byte DisplayIssueDateFlag { get; set; }
        public byte DisplayIssueNumberFlag { get; set; }

        public string FormOfPaymentSubtypeRcd { get; set; }
        public string DisplayName { get; set; }
        public string FormOfPaymentRcd { get; set; }
        public string CardCode { get; set; }



        #endregion
    }
}
