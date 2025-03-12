using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

namespace Avantik.Web.Service.Entity.Client
{
    public class ClientProfile 
    {
        public Client Client { get; set; }
        public IList<PassengerProfile> PassengerProfiles { get; set; }
        public IList<Entity.Booking.Remark> BookingRemarks { get; set; }

    }
}
