using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.ServiceModel;
using System.Text;
using Avantik.Web.Service.Message.ManageBooking;

namespace Avantik.Web.Service.Message.ManageBooking
{
    [MessageContract]
    public class SegmentCancelResponse : ResponseBase
    {
    }
}
