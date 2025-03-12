using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.Serialization;
using System.ServiceModel;

namespace Avantik.Web.Service.Message.Booking
{
    public class PassengerSegmentMapping 
    {
        #region Property
        [MessageBodyMember]
        public string BookingSegmentId { get; set; }
        [MessageBodyMember]
        public string PassengerId { get; set; }

        #endregion
    }
}
