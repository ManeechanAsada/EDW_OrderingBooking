using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.Serialization;
using System.ServiceModel;

namespace Avantik.Web.Service.Message.Fee
{
    public class Fee
    {
        [MessageBodyMember]
        public Guid AccountFeeBy { get; set; }
        [DataMember(EmitDefaultValue = false)]
        public DateTime AccountFeeDateTime { get; set; }
        [MessageBodyMember]
        public decimal AcctFeeAmount { get; set; }
        [MessageBodyMember]
        public decimal AcctFeeAmountIncl { get; set; }
        [MessageBodyMember]
        public string AgencyCode { get; set; }
        [MessageBodyMember]
        public Guid BaggageFeeOptionId { get; set; }
        [MessageBodyMember]
        public Guid BookingFeeId { get; set; }
        [MessageBodyMember]
        public Guid BookingId { get; set; }
        [MessageBodyMember]
        public Guid BookingSegmentId { get; set; }
        [MessageBodyMember]
        public string ChangeComment { get; set; }
        [MessageBodyMember]
        public decimal ChargeAmount { get; set; }
        [MessageBodyMember]
        public decimal ChargeAmountIncl { get; set; }
        [MessageBodyMember]
        public string ChargeCurrencyRcd { get; set; }
        [MessageBodyMember]
        public string Comment { get; set; }
        [MessageBodyMember]
        public Guid CreateBy { get; set; }
        [DataMember(EmitDefaultValue = false)]
        public DateTime CreateDateTime { get; set; }
        [MessageBodyMember]
        public string CurrencyRcd { get; set; }
        [MessageBodyMember]
        public string DestinationRcd { get; set; }
        [MessageBodyMember]
        public string DisplayName { get; set; }
        [DataMember(EmitDefaultValue = false)]
        public DateTime DocumentDateTime { get; set; }
        [MessageBodyMember]
        public string DocumentNumber { get; set; }
        [MessageBodyMember]
        public string ExternalReference { get; set; }
        [MessageBodyMember]
        public decimal FeeAmount { get; set; }
        [MessageBodyMember]
        public decimal FeeAmountIncl { get; set; }
        [MessageBodyMember]
        public string FeeCalculationRcd { get; set; }
        [MessageBodyMember]
        public string FeeCategoryRcd { get; set; }
        [MessageBodyMember]
        public Guid FeeId { get; set; }
        [MessageBodyMember]
        public string FeeLevel { get; set; }
        [MessageBodyMember]
        public decimal FeePercentage { get; set; }
        [MessageBodyMember]
        public string FeeRcd { get; set; }
        [MessageBodyMember]
        public byte ManualFeeFlag { get; set; }
        [MessageBodyMember]
        public byte MinimumFeeAmountFlag { get; set; }
        [MessageBodyMember]
        public bool NewRecord { get; set; }
        [MessageBodyMember]
        public decimal NumberOfUnits { get; set; }
        [MessageBodyMember]
        public string OdDestinationRcd { get; set; }
        [MessageBodyMember]
        public byte OdFlag { get; set; }
        [MessageBodyMember]
        public string OdOriginRcd { get; set; }
        [MessageBodyMember]
        public string OriginRcd { get; set; }
        [MessageBodyMember]
        public Guid PassengerId { get; set; }
        [MessageBodyMember]
        public Guid PassengerSegmentServiceId { get; set; }
        [MessageBodyMember]
        public decimal PaymentAmount { get; set; }
        [MessageBodyMember]
        public bool SelectedFee { get; set; }
        [MessageBodyMember]
        public byte SkipFareAllowanceFlag { get; set; }
        [MessageBodyMember]
        public decimal TotalAmount { get; set; }
        [MessageBodyMember]
        public decimal TotalAmountIncl { get; set; }
        [MessageBodyMember]
        public decimal TotalFeeAmount { get; set; }
        [MessageBodyMember]
        public decimal TotalFeeAmountIncl { get; set; }
        [MessageBodyMember]
        public string Units { get; set; }
        [MessageBodyMember]
        public Guid UpdateBy { get; set; }
        [DataMember(EmitDefaultValue = false)]
        public DateTime UpdateDateTime { get; set; }
        [MessageBodyMember]
        public decimal VatPercentage { get; set; }
        [MessageBodyMember]
        public string VendorRcd { get; set; }
        [MessageBodyMember]
        public Guid VoidBy { get; set; }
        [DataMember(EmitDefaultValue = false)]
        public DateTime VoidDateTime { get; set; }
    }
}
