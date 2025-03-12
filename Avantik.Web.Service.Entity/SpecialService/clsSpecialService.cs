using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

namespace Avantik.Web.Service.Entity
{
    public class SpecialService
    {
        public Guid BookingSegmentId { get; set; }
        public Guid CreateBy { get; set; }
        public DateTime CreateDateTime { get; set; }
        public int CutOffTime { get; set; }
        public string DestinationRcd { get; set; }
        public string DisplayName { get; set; }
        public string FeeId { get; set; }
        public string FlightId { get; set; }
        public string HelpText { get; set; }
        public byte IncludeActionCodeFlag { get; set; }
        public byte IncludeCateringFlag { get; set; }
        public byte IncludeFlightSegmentFlag { get; set; }
        public byte IncludeNumberOfServiceFlag { get; set; }
        public byte IncludePassengerAssistanceFlag { get; set; }
        public byte IncludePassengerNameFlag { get; set; }
        public byte InventoryControlFlag { get; set; }
        public byte ManifestFlag { get; set; }
        public bool NewRecord { get; set; }
        public short NumberOfUnits { get; set; }
        public string OriginRcd { get; set; }
        public Guid PassengerId { get; set; }
        public Guid PassengerSegmentServiceId { get; set; }
        public byte SendInterlineReplyFlag { get; set; }
        public byte ServiceOnRequestFlag { get; set; }
        public byte ServiceSupportedFlag { get; set; }
        public string ServiceText { get; set; }
        public string SpecialServiceChangeStatusRcd { get; set; }
        public string SpecialServiceGroupRcd { get; set; }
        public string SpecialServiceRcd { get; set; }
        public string SpecialServiceStatusRcd { get; set; }
        public string StatusCode { get; set; }
        public byte TextAllowedFlag { get; set; }
        public byte TextRequiredFlag { get; set; }
        public Guid UpdateBy { get; set; }
        public DateTime UpdateDateTime { get; set; }

    }
}
