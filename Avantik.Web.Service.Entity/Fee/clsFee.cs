using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

namespace Avantik.Web.Service.Entity
{
    public class Fee
    {
        public Guid AccountFeeBy { get; set; }
        public DateTime AccountFeeDateTime { get; set; }
        public decimal AcctFeeAmount { get; set; }
        public decimal AcctFeeAmountIncl { get; set; }
        public string AgencyCode { get; set; }
        public Guid BaggageFeeOptionId { get; set; }
        public Guid BookingFeeId { get; set; }
        public Guid BookingId { get; set; }
        public Guid BookingSegmentId { get; set; }
        public string ChangeComment { get; set; }
        public decimal ChargeAmount { get; set; }
        public decimal ChargeAmountIncl { get; set; }
        public string ChargeCurrencyRcd { get; set; }
        public string Comment { get; set; }
        public Guid CreateBy { get; set; }
        public DateTime CreateDateTime { get; set; }
        public string CurrencyRcd { get; set; }
        public string DestinationRcd { get; set; }
        public string DisplayName { get; set; }
        public DateTime DocumentDateTime { get; set; }
        public string DocumentNumber { get; set; }
        public string ExternalReference { get; set; }
        public decimal FeeAmount { get; set; }
        public decimal FeeAmountIncl { get; set; }
        public string FeeCalculationRcd { get; set; }
        public string FeeCategoryRcd { get; set; }
        public Guid FeeId { get; set; }
        public string FeeLevel { get; set; }
        public decimal FeePercentage { get; set; }
        public string FeeRcd { get; set; }
        public byte ManualFeeFlag { get; set; }
        public byte MinimumFeeAmountFlag { get; set; }
        public bool NewRecord { get; set; }
        public decimal NumberOfUnits { get; set; }
        public string OdDestinationRcd { get; set; }
        public byte OdFlag { get; set; }
        public string OdOriginRcd { get; set; }
        public string OriginRcd { get; set; }
        public Guid PassengerId { get; set; }
        public Guid PassengerSegmentServiceId { get; set; }
        public decimal PaymentAmount { get; set; }
        public bool SelectedFee { get; set; }
        public byte SkipFareAllowanceFlag { get; set; }
        public decimal TotalAmount { get; set; }
        public decimal TotalAmountIncl { get; set; }
        public decimal TotalFeeAmount { get; set; }
        public decimal TotalFeeAmountIncl { get; set; }
        public string units { get; set; }
        public Guid UpdateBy { get; set; }
        public DateTime UpdateDateTime { get; set; }
        public decimal VatPercentage { get; set; }
        public string VendorRcd { get; set; }
        public Guid VoidBy { get; set; }
        public DateTime VoidDateTime { get; set; }
    }
}
