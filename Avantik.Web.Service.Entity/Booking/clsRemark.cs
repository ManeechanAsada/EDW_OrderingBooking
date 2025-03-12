using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Avantik.Web.Service.Entity.Booking

{
    public class Remark
    {
        public Guid BookingRemarkId { get; set; }
        public string RemarkTypeRcd { get; set; }
        public DateTime TimelimitDateTime { get; set; }
        public byte CompleteFlag { get; set; }
        public string RemarkText { get; set; }
        public Guid BookingId { get; set; }
        public string AddedBy { get; set; }
        public Guid ClientProfileId { get; set; }
        public string AgencyCode { get; set; }
        public Guid CreateBy { get; set; }
        public DateTime CreateDateTime { get; set; }
        public Guid UpdateBy { get; set; }
        public DateTime UpdateDateTime { get; set; }
        public byte SystemFlag { get; set; }
        public DateTime UtcTimelimitDateTime { get; set; }
        public string Nickname { get; set; }
        public byte ProtectedFlag { get; set; }
        public byte WarningFlag { get; set; }
        public byte ProcessMessageFlag { get; set; }
       
    } 

}
