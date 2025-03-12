using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Avantik.Web.Service.Entity.Flight
{
    public class Availabilities
    {
        public IEnumerable<Availability> FlightAvailabilityOutbound { get; set; }
        public IEnumerable<Availability> FlightAvailabilityReturn { get; set; }
    }
}
