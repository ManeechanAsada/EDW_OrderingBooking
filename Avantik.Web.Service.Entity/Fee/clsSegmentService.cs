using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

namespace Avantik.Web.Service.Entity
{
    public class SegmentService
    {
        public Guid FlightConnectionId { get; set; }
        public string SpecialServiceRcd { get; set; }
        public string OriginRcd { get; set; }
        public string DestinationRcd { get; set; }
        public string OdOriginRcd { get; set; }
        public string OdDestinationRcd { get; set; }
        public string BookingClassRcd { get; set; }
        public string FareCode { get; set; }
        public string AirlineRcd { get; set; }
        public string FlightNumber { get; set; }
        public DateTime DepartureDate { get; set; }
        public Guid BookingSegmentId { get; set; }
    }
}
