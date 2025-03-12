using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Avantik.Web.Service.Entity.Flight
{
    public static class AvailabilityExtension
    {
        public static bool FindAvlNetTotalFlag(this IList<Availability> availability)
        {
            for (int i = 0; i < availability.Count; i++)
            {
                if (availability[i].avl_show_net_total_flag == 1)
                {
                    return true;
                }
            }
            return false;
        }
    }
}
