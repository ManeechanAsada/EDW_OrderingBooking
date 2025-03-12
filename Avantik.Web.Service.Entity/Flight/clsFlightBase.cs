using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

namespace Avantik.Web.Service.Entity.Flight
{
    public abstract class FlightBase
    {
        public string airline_rcd { get; set; }
        public string flight_number { get; set; }
        public string flight_status_rcd { get; set; }
        public string aircraft_type_rcd { get; set; }
        public string origin_rcd { get; set; }
        public string destination_rcd { get; set; }
        public string flight_comment { get; set; }
    }
}
