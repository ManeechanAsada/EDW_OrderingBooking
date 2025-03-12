using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using Avantik.Web.Service.Entity.Flight;

namespace Avantik.Web.Service.Model
{
    public interface IAvailabilityBase
    {
        IList<Availability> GetAvailability();
    }
}
