using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Avantik.Web.Service.Entity.Route
{
    public abstract class RouteBase
    {
        public string origin_rcd { get; set; }
        public string destination_rcd { get; set; }

        public byte segment_change_fee_flag { get; set; }
        public byte special_service_fee_flag { get; set; }
        public byte show_insurance_on_web_flag { get; set; }
    }
}
