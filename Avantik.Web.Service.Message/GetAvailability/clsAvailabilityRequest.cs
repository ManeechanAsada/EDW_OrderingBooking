using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using System.ServiceModel;
using System.Runtime.Serialization;
using Avantik.Web.Service;

namespace Avantik.Web.Service.Message
{
    [MessageContract]
    public class AvailabilityRequest
    {
        [MessageHeader]
        public string Token { get; set; }
        [MessageBodyMember]
        public string OriginRcd { get; set; }
        [MessageBodyMember]
        public string DestinationRcd { get; set; }
        [MessageBodyMember]
        public string OdOriginRcd { get; set; }
        [MessageBodyMember]
        public string OdDestinationRcd { get; set; }
        [MessageBodyMember]
        public string OtherPassengerType { get; set; }
        [MessageBodyMember]
        public string BoardingClass { get; set; }
        [MessageBodyMember]
        public string BookingClass { get; set; }
        [MessageBodyMember]
        public string DayTimeIndicator { get; set; }
        [MessageBodyMember]
        public string AgencyCode { get; set; }
        [MessageBodyMember]
        public string CurrencyCode { get; set; }
        [MessageBodyMember]
        public string ReturnDayTimeIndicator { get; set; }
        [MessageBodyMember]
        public string PromotionCode { get; set; }
        [MessageBodyMember]
        public Infrastructure.AvailabilityFareTypes FareTypes { get; set; }
        [MessageBodyMember]
        public string LanguageCode { get; set; }
        [MessageBodyMember]
        public string IpAddress { get; set; }
        [MessageBodyMember]
        public string Transitpoint { get; set; }
        [MessageBodyMember]
        public string FlightId { get; set; }
        [MessageBodyMember]
        public string FareId { get; set; }
        [MessageBodyMember]
        public string TransactionReference { get; set; }


        [MessageBodyMember]
        public DateTime FromDate { get; set; }
        [MessageBodyMember]
        public DateTime ToDate { get; set; }
        [MessageBodyMember]
        public DateTime ReturnFromDate { get; set; }
        [MessageBodyMember]
        public DateTime ReturnToDate { get; set; }
        [MessageBodyMember]
        public DateTime BookingDate { get; set; }


        [MessageBodyMember]
        public decimal MaxAmount { get; set; }

        [MessageBodyMember]
        public bool ShowClose { get; set; }
        [MessageBodyMember]
        public bool MapWithFares { get; set; }
        [MessageBodyMember]
        public bool ApplyFareLogic { get; set; }
        [MessageBodyMember]
        public Int16 NoneStopOnly { get; set; }
        
        [MessageBodyMember]
        public byte Adult { get; set; }
        [MessageBodyMember]
        public byte Child { get; set; }
        [MessageBodyMember]
        public byte Infant { get; set; }
        [MessageBodyMember]
        public byte Other { get; set; }
        [MessageBodyMember]
        public Int16 IncludeDeparted { get; set; }
        [MessageBodyMember]
        public Int16 IncludeCancelled { get; set; }
        [MessageBodyMember]
        public Int16 IncludeWaitlisted { get; set; }
        [MessageBodyMember]
        public Int16 IncludeSoldOut { get; set; }
        [MessageBodyMember]
        public Int16 IncludeFares { get; set; }
        [MessageBodyMember]
        public Int16 Refundable { get; set; }
        [MessageBodyMember]
        public Int16 ReturnRefundable { get; set; }
        [MessageBodyMember]
        public Int16 GroupFares { get; set; }
        [MessageBodyMember]
        public Int16 ITFaresOnly { get; set; }
        [MessageBodyMember]
        public Int16 StaffFares { get; set; }
        [MessageBodyMember]
        public bool NoVat { get; set; }
        
        [MessageBodyMember]
        public Infrastructure.AvailabilityTypes AvailabilityTypes { get; set; }
    }
}
