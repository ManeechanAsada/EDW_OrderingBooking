using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using System.ServiceModel;
using System.Runtime.Serialization;
using Avantik.Web.Service;
using Avantik.Web.Service.Message;

namespace Avantik.Web.Service.Message
{
      [MessageContract]
    public class CalculateFeesBookingResponse : ResponseBase
    {
          [MessageBodyMember]
          public IList<Message.Booking.Fee> Fees { get; set; }

    }
}
