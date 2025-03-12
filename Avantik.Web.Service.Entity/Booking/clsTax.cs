using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Avantik.Web.Service.Entity.Booking

{
    public class Tax
    {
        public Guid PassengerSegmentTaxId { get; set; }
        public decimal TaxAmount { get; set; }
        public decimal TaxAmountIncl { get; set; }
        public decimal AcctAmount { get; set; }
        public decimal AcctAmountIncl { get; set; }
        public decimal SalesAmount { get; set; }
        public decimal SalesAmountIncl { get; set; }
        public string TaxRcd { get; set; }
        public Guid PassengerId { get; set; }
        public Guid TaxId { get; set; }
        public Guid BookingSegmentId { get; set; }
        public string TaxCurrencyRcd { get; set; }
        public string SalesCurrencyRcd { get; set; }
        public string DisplayName { get; set; }
        public string SummarizeUp { get; set; }
        public string CoverageType { get; set; }
        public Guid CreateBy { get; set; }
        public Guid UpdateBy { get; set; }
        public DateTime CreateDateTime { get; set; }
        public DateTime UpdateDateTime { get; set; }
        public decimal VatPercentage { get; set; }
    }
}
