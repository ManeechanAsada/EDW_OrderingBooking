using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Avantik.Web.Service.Entity.Booking
{
    public class Flight : FlightBase
    {
        #region property
        public Guid FairId { get; set; }
        public Guid FlightConnectionId { get; set; }
        public Guid ExchangedSegmentId { get; set; }
        public string BoardingClassRcd { get; set; }
        public string BookingClassRcd { get; set; }
        public string TransitPoints { get; set; }
        public string TransitpointsName { get; set; }
        public int NumberOfUnits { get; set; }
        public byte EticketFlag { get; set; }

        
        #endregion
    }
}
