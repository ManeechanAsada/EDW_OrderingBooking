using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.Serialization;
using System.ServiceModel;

namespace Avantik.Web.Service.Message.SSR
{
    public class Service 
    {
 	 	 [MessageBodyMember]
        public string PassengerId { get; set; }
 	 	 [MessageBodyMember]
         public string BookingSegmentId { get; set; }
 	 	 [MessageBodyMember] 
        public string SpecialServiceCode  { get; set; }
         [MessageBodyMember] 
        public Int16 NumberOfUnits  { get; set; }
    }
}
