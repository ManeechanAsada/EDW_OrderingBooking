using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Avantik.Web.Service.Entity
{
    public abstract class FlightBase
    {
        public Guid FlightId { get; set; }
        public string OriginRcd { get; set; }
        public string DestinationRcd { get; set; }
        public string OdOriginRcd { get; set; }
        public string OdDestinationRcd { get; set; }
        public string AirlineRcd { get; set; }
        public string FlightNumber { get; set; }
        public DateTime DepartureDate { get; set; }
    }
}
