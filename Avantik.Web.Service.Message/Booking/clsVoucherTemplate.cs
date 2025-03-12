using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.Serialization;
using System.ServiceModel;

namespace Avantik.Web.Service.Message
{
    public class VoucherTemplate
    {
        [MessageBodyMember]
        public byte AirlineFlag { get; set; }
        [MessageBodyMember]
        public byte B2bFlag { get; set; }
        [MessageBodyMember]
        public byte B2cFlag { get; set; }
        [MessageBodyMember]
        public byte B2eFlag { get; set; }
        [MessageBodyMember]
        public decimal ChargeAmount { get; set; }
        [MessageBodyMember]
        public string CurrencyRcd { get; set; }
        [MessageBodyMember]
        public string Destinations { get; set; }
        [MessageBodyMember]
        public decimal DiscountPercentage { get; set; }
        [MessageBodyMember]
        public string DisplayName { get; set; }
        [MessageBodyMember]
        public byte FareOnlyFlag { get; set; }
        [MessageBodyMember]
        public string FormOfPaymentRcd { get; set; }
        [MessageBodyMember]
        public string FormOfPaymentSubtypeRcd { get; set; }
        [MessageBodyMember]
        public byte MultiplePaymentFlag { get; set; }
        [MessageBodyMember]
        public string Origins { get; set; }
        [MessageBodyMember]
        public byte OtherFeeFlag { get; set; }
        [MessageBodyMember]
        public short PassengerSegments { get; set; }
        [MessageBodyMember]
        public byte RecipientOnlyFlag { get; set; }
        [MessageBodyMember]
        public byte SeatFeeFlag { get; set; }
        [MessageBodyMember]
        public string StatusCode { get; set; }
        [MessageBodyMember]
        public byte TicketFlag { get; set; }
        [MessageBodyMember]
        public int ValidDays { get; set; }
        [MessageBodyMember]
        public string ValidForClass { get; set; }
        [MessageBodyMember]
        public DateTime ValidFromDate { get; set; }
        [MessageBodyMember]
        public DateTime ValidToDate { get; set; }
        [MessageBodyMember]
        public Guid VoucherTemplateId { get; set; }
        [MessageBodyMember]
        public string VoucherUseCode { get; set; }
        [MessageBodyMember]
        public decimal VoucherValue { get; set; }
    }
}
