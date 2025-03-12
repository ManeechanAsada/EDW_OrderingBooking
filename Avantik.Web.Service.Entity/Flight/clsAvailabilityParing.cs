using System;
using System.Collections.Generic;
using System.Text;

namespace Avantik.Web.Service.Entity.Flight
{
    public class AvailabilityParing : AvailabilityRoute
    {
        public string leg_1_airline_rcd { get; set; }
        public string leg_1_flight_number { get; set; }
        public string leg_2_airline_rcd { get; set; }
        public string leg_2_flight_number { get; set; }
    }
}
