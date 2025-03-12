using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Avantik.Web.Service.Infrastructure
{
    public enum AvailabilityTypes
    {
        OWN,
        TOUROPERATOR
    };

    public enum AvailabilityFareTypes
    {
        ALL,
        LOWEST,
        LOWESTGROUP,
        LOWESTCLASS,
        REDEMPTION
    };
}
