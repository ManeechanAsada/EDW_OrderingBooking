using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ServiceModel;
using System.Runtime.Serialization;
using Avantik.Web.Service.Message.Booking;
using Avantik.Web.Service.Contracts;
using Avantik.Web.Service.Message;
using Avantik.Web.Service.Message.Fee;

namespace Avantik.Web.Service.Proxy
{
    public class AvailabilityServiceProxy : ClientBase<IAvailabilityService>, IAvailabilityService
    {
        //public AvailabilityResponse GetAvailability(AvailabilityRequest request)
        //{
        //    return base.Channel.GetAvailability(request);
        //}
    }
}
