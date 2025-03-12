using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.Serialization;
using System.ServiceModel;

namespace Avantik.Web.Service.Message.Booking
{
    public class Tax 
    {
        [MessageBodyMember]
        public Guid PassengerSegmentTaxId { get; set; }
        [MessageBodyMember]
        public decimal TaxAmount { get; set; }
        [MessageBodyMember]
        public decimal TaxAmountIncl { get; set; }
        [MessageBodyMember]
        public decimal AcctAmount { get; set; }
        [MessageBodyMember]
        public decimal AcctAmountIncl { get; set; }
        [MessageBodyMember]
        public decimal SalesAmount { get; set; }
        [MessageBodyMember]
        public decimal SalesAmountIncl { get; set; }
        [MessageBodyMember]
        public string TaxRcd { get; set; }
        [MessageBodyMember]
        public Guid PassengerId { get; set; }
        [MessageBodyMember]
        public Guid TaxId { get; set; }
        [MessageBodyMember]
        public Guid BookingSegmentId { get; set; }
        [MessageBodyMember]
        public string TaxCurrencyRcd { get; set; }
        [MessageBodyMember]
        public string SalesCurrencyRcd { get; set; }
        [MessageBodyMember]
        public string DisplayName { get; set; }
        [MessageBodyMember]
        public string SummarizeUp { get; set; }
        [MessageBodyMember]
        public string CoverageType { get; set; }
        [MessageBodyMember]
        public Guid CreateBy { get; set; }
        [MessageBodyMember]
        public Guid UpdateBy { get; set; }
        [MessageBodyMember]
        public DateTime CreateDateTime { get; set; }
        [MessageBodyMember]
        public DateTime UpdateDateTime { get; set; }
        [MessageBodyMember]
        public decimal VatPercentage { get; set; }
    }
}
