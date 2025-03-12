using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Avantik.Web.Service.Entity.Booking
{
    public class Fee
    {
        public Guid BookingFeeId  { get; set; }
        public decimal FeeAmount  { get; set; }
        public Guid BookingId  { get; set; }
        public Guid PassengerId  { get; set; }
        public string CurrencyRcd  { get; set; }
        public decimal AcctFeeAmount  { get; set; }
        public Guid FeeId  { get; set; }
        public decimal VatPercentage  { get; set; }
        public decimal FeeAmountIncl  { get; set; }
        public decimal AcctFeeAmountIncl  { get; set; }
        public string FeeRcd  { get; set; }
        public string DisplayName  { get; set; }
        public Guid AccountFeeBy  { get; set; }
        public DateTime AccountFeeDateTime  { get; set; }
        public DateTime VoidDateTime  { get; set; }
        public Guid VoidBy  { get; set; }
        public decimal PaymentAmount  { get; set; }
        public Guid CreateBy  { get; set; }
        public DateTime CreateDateTime  { get; set; }
        public Guid UpdateBy  { get; set; }
        public DateTime UpdateDateTime  { get; set; }
        public Guid BookingSegmentId  { get; set; }
        public string AgencyCode  { get; set; }
        public Guid PassengerSegmentServiceId  { get; set; }
        public string FeeCategoryRcd  { get; set; }
        public string OriginRcd  { get; set; }
        public string DestinationRcd  { get; set; }
        public string OdOriginRcd  { get; set; }
        public string OdDestinationRcd  { get; set; }
        public decimal NumberOfUnits  { get; set; }
        public decimal TotalAmount  { get; set; }
        public decimal TotalAmountIncl  { get; set; }
        public byte ManualFeeFlag  { get; set; }
        public byte OdFlag  { get; set; }
        public byte SkipFareAllowanceFlag  { get; set; }
        public string FeeLevel  { get; set; }
        public string FeeCalculationRcd  { get; set; }
        public byte MinimumFeeAmountFlag  { get; set; }
        public decimal FeePercentage  { get; set; }
        public string ChangeComment  { get; set; }
        public string Comment  { get; set; }
        public decimal TotalFeeAmount  { get; set; }
        public decimal TotalFeeAmountIncl  { get; set; }
        public string Units  { get; set; }
        public string ChargeCurrencyRcd  { get; set; }
        public decimal ChargeAmount  { get; set; }
        public decimal ChargeAmountIncl  { get; set; }
        public string DocumentNumber  { get; set; }
        public DateTime DocumentDateTime  { get; set; }
        public string ExternalReference  { get; set; }
        public string VendorRcd  { get; set; }
        public Guid BaggageFeeOptionId  { get; set; }
        public bool NewRecord  { get; set; }
        public bool SelectedFee  { get; set; }
        public string UnitRcd { get; set; }
        public decimal WeightLbs { get; set; }
        public decimal WeightKgs { get; set; }
        public string MpdNumber { get; set; }

        #region Method
        public PaymentAllocation GetAllocation(Guid bookingPaymentId, Guid userId)
        {
            if (FeeAmountIncl - PaymentAmount > 0)
            {
                PaymentAllocation allocation = new PaymentAllocation();

                allocation.booking_payment_id = bookingPaymentId;
                allocation.passenger_id = PassengerId;
                allocation.booking_segment_id = BookingSegmentId;
                allocation.booking_fee_id = BookingFeeId;
                allocation.passenger_segment_service_id = PassengerSegmentServiceId;
                //allocation.fee_id = FeeId;
                allocation.user_id = userId;

                allocation.currency_rcd = CurrencyRcd;
                allocation.charge_currency_rcd = ChargeCurrencyRcd;
                allocation.od_origin_rcd = OdOriginRcd;
                allocation.od_destination_rcd = OdDestinationRcd;
                allocation.fee_category_rcd = FeeCategoryRcd;
                allocation.vendor_rcd = VendorRcd;
                allocation.units = Units;
                allocation.external_reference = ExternalReference;

                allocation.sales_amount = FeeAmountIncl - PaymentAmount;
                allocation.payment_amount = PaymentAmount;
                allocation.weight_lbs = WeightLbs;
                allocation.weight_kgs = WeightKgs;

                return allocation;
            }
            else
            {
                return null;
            }
        }
        #endregion
    }
}
