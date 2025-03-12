using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Avantik.Web.Service.Entity.Booking
{
    public class PassengerService
    {
        public Guid PassengerSegmentServiceId { get; set; }
        public Guid PassengerId { get; set; }
        public Guid BookingSegmentId { get; set; }
        public string SpecialServiceStatusRcd { get; set; }
        public string SpecialServiceChangeStatusRcd { get; set; }
        public string SpecialServiceRcd { get; set; }
        public string ServiceText { get; set; }
        public Guid CreateBy { get; set; }
        public DateTime CreateDateTime { get; set; }
        public Guid UpdateBy { get; set; }
        public DateTime UpdateDateTime { get; set; }
        public string FlightId { get; set; }
        public string FeeId { get; set; }
        public Int16 NumberOfUnits { get; set; }
        public string OriginRcd { get; set; }
        public string DestinationRcd { get; set; }
        public bool NewRecord { get; set; }
        public string DisplayName { get; set; }
        public int CutOffTime { get; set; }
        public string StatusCode { get; set; }
        public string HelpText { get; set; }
        public string SpecialServiceGroupRcd { get; set; }
        public byte ServiceOnRequestFlag { get; set; }
        public byte TextAllowedFlag { get; set; }
        public byte ManifestFlag { get; set; }
        public byte TextRequiredFlag { get; set; }
        public byte IncludePassengerNameFlag { get; set; }
        public byte IncludeFlightSegmentFlag { get; set; }
        public byte IncludeActionCodeFlag { get; set; }
        public byte IncludeNumberOfServiceFlag { get; set; }
        public byte IncludeCateringFlag { get; set; }
        public byte IncludePassengerAssistanceFlag { get; set; }
        public byte ServiceSupportedFlag { get; set; }
        public byte SendInterlineReplyFlag { get; set; }
        public byte InventoryControlFlag { get; set; }
        public string UnitRcd { get; set; }

    }
}
