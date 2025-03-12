using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Avantik.Web.Service.Entity.Fares
{
    public class FareLogic
    {
        public int level_1_prior_days { get; set; }
        public int level_2_prior_days { get; set; }
        public int level_2_1_days { get; set; }
        public int level_2_2_days { get; set; }
        public int level_3_1_days { get; set; }
        public int level_3_2_days { get; set; }

        public string level_1_oneway_classes { get; set; }
        public string level_1_return_classes { get; set; }
        public string level_2_oneway_classes { get; set; }
        public string level_2_1_return_classes { get; set; }
        public string level_2_2_shortstay_classes { get; set; }
        public string level_2_2_return_classes { get; set; }
        public string level_3_oneway_classes { get; set; }
        public string level_3_1_return_classes { get; set; }
        public string level_3_2_shortstay_classes { get; set; }
        public string level_3_2_return_classes { get; set; }
    }
}
