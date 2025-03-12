using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using Avantik.Web.Service.Entity.Flight;
using Avantik.Web.Service.Repository;

namespace Avantik.Web.Service.Model
{
    public abstract class AvailabilityDecorator : IAvailabilityBase
    {
        protected IAvailabilityBase _Availability;

        public AvailabilityDecorator(IAvailabilityBase availability)
        {
            _Availability = availability;
        }

        public abstract IList<Availability> GetAvailability();
    }
}
