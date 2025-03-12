using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Avantik.Web.Service.Entity.Flight
{
    public class AvailabilityRoute : Route.RouteBase
    {
        public string transit_airport_rcd { get; set; }
        public string avl_origin_rcd { get; set; }
        public string avl_destination_rcd { get; set; }

        public byte transit_flag { get; set; }
        public byte direct_flag { get; set; }
        public byte dynamic_connections_flag { get; set; }
        
        public int min_transit_minutes { get; set; }
        public int max_transit_minutes { get; set; }
        public int connex { get; set; }
    }
}
