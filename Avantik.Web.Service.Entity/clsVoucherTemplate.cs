using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Avantik.Web.Service.Entity
{
    public class VoucherTemplate
    {
        public byte AirlineFlag { get; set; }
        public byte B2bFlag { get; set; }
        public byte B2cFlag { get; set; }
        public byte B2eFlag { get; set; }
        public decimal ChargeAmount { get; set; }
        public string CurrencyRcd { get; set; }
        public string Destinations { get; set; }
        public decimal DiscountPercentage { get; set; }
        public string DisplayName { get; set; }
        public byte FareOnlyFlag { get; set; }
        public string FormOfPaymentRcd { get; set; }
        public string FormOfPaymentSubtypeRcd { get; set; }
        public byte MultiplePaymentFlag { get; set; }
        public string Origins { get; set; }
        public byte OtherFeeFlag { get; set; }
        public short PassengerSegments { get; set; }
        public byte RecipientOnlyFlag { get; set; }
        public byte SeatFeeFlag { get; set; }
        public string StatusCode { get; set; }
        public byte TicketFlag { get; set; }
        public int ValidDays { get; set; }
        public string ValidForClass { get; set; }
        public DateTime ValidFromDate { get; set; }
        public DateTime ValidToDate { get; set; }
        public Guid VoucherTemplateId { get; set; }
        public string VoucherUseCode { get; set; }
        public decimal VoucherValue { get; set; }
    }
}
