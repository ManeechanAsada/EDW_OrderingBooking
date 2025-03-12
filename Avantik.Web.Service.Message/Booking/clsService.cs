using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.Serialization;
using System.ServiceModel;

namespace Avantik.Web.Service.Message.Booking
{
    public class PassengerService 
    {
        [MessageBodyMember]
        public Guid PassengerSegmentServiceId { get; set; }
 	 	 [MessageBodyMember] 
        public Guid PassengerId  { get; set; }
 	 	 [MessageBodyMember] 
        public Guid BookingSegmentId  { get; set; }
 	 	 [MessageBodyMember] 
        public string SpecialServiceStatusRcd  { get; set; }
 	 	 [MessageBodyMember] 
        public string SpecialServiceChangeStatusRcd  { get; set; }
 	 	 [MessageBodyMember] 
        public string SpecialServiceRcd  { get; set; }
 	 	 [MessageBodyMember] 
        public string ServiceText  { get; set; }
 	 	 [MessageBodyMember] 
        public Guid CreateBy  { get; set; }
         [MessageBodyMember]
        public DateTime CreateDateTime  { get; set; }
 	 	 [MessageBodyMember] 
        public Guid UpdateBy  { get; set; }
         [MessageBodyMember]
        public DateTime UpdateDateTime  { get; set; }
 	 	 [MessageBodyMember] 
        public string FlightId  { get; set; }
 	 	 [MessageBodyMember] 
        public string FeeId  { get; set; }
 	 	 [MessageBodyMember] 
        public Int16 NumberOfUnits  { get; set; }
 	 	 [MessageBodyMember] 
        public string OriginRcd  { get; set; }
 	 	 [MessageBodyMember] 
        public string DestinationRcd  { get; set; }
 	 	 [MessageBodyMember] 
        public bool NewRecord  { get; set; }
 	 	 [MessageBodyMember] 
        public string DisplayName  { get; set; }
 	 	 [MessageBodyMember] 
        public int CutOffTime  { get; set; }
 	 	 [MessageBodyMember] 
        public string StatusCode  { get; set; }
 	 	 [MessageBodyMember] 
        public string HelpText  { get; set; }
 	 	 [MessageBodyMember] 
        public string SpecialServiceGroupRcd  { get; set; }
 	 	 [MessageBodyMember] 
        public byte ServiceOnRequestFlag  { get; set; }
 	 	 [MessageBodyMember] 
        public byte TextAllowedFlag  { get; set; }
 	 	 [MessageBodyMember] 
        public byte ManifestFlag  { get; set; }
 	 	 [MessageBodyMember] 
        public byte TextRequiredFlag  { get; set; }
 	 	 [MessageBodyMember] 
        public byte IncludePassengerNameFlag  { get; set; }
 	 	 [MessageBodyMember] 
        public byte IncludeFlightSegmentFlag  { get; set; }
 	 	 [MessageBodyMember] 
        public byte IncludeActionCodeFlag  { get; set; }
 	 	 [MessageBodyMember] 
        public byte IncludeNumberOfServiceFlag  { get; set; }
 	 	 [MessageBodyMember] 
        public byte IncludeCateringFlag  { get; set; }
 	 	 [MessageBodyMember] 
        public byte IncludePassengerAssistanceFlag  { get; set; }
 	 	 [MessageBodyMember] 
        public byte ServiceSupportedFlag  { get; set; }
 	 	 [MessageBodyMember] 
        public byte SendInterlineReplyFlag  { get; set; }
 	 	 [MessageBodyMember] 
        public byte InventoryControlFlag  { get; set; }
         [MessageBodyMember]
         public string UnitRcd { get; set; }

    }
}
