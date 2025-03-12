using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.Serialization;
using System.ServiceModel;

namespace Avantik.Web.Service.Message.Booking
{
    public class Remark 
    {
        [MessageBodyMember]
        public Guid BookingRemarkId { get; set; }
        [MessageBodyMember]
        public string RemarkTypeRcd { get; set; }
        [MessageBodyMember]
        public DateTime TimelimitDateTime { get; set; }
        [MessageBodyMember]
        public byte CompleteFlag { get; set; }
        [MessageBodyMember]
        public string RemarkText { get; set; }
        [MessageBodyMember]
        public Guid BookingId { get; set; }
        [MessageBodyMember]
        public string AddedBy { get; set; }
        [MessageBodyMember]
        public Guid ClientProfileId { get; set; }
        [MessageBodyMember]
        public string AgencyCode { get; set; }
        [MessageBodyMember]
        public Guid CreateBy { get; set; }
        [MessageBodyMember]
        public DateTime CreateDateTime { get; set; }
        [MessageBodyMember]
        public Guid UpdateBy { get; set; }
        [MessageBodyMember]
        public DateTime UpdateDateTime { get; set; }
        [MessageBodyMember]
        public byte SystemFlag { get; set; }
        [MessageBodyMember]
        public DateTime UtcTimelimitDateTime { get; set; }
        [MessageBodyMember]
        public string Nickname { get; set; }
        [MessageBodyMember]
        public byte ProtectedFlag { get; set; }
        [MessageBodyMember]
        public byte WarningFlag { get; set; }
        [MessageBodyMember]
        public byte ProcessMessageFlag { get; set; }
    }
}
