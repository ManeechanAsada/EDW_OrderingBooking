using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Avantik.Web.Service.Entity.Flight;

namespace Avantik.Web.Service.Model
{
    public class AvailabilityLowestClass : AvailabilityDecorator
    {
        public AvailabilityLowestClass(IAvailabilityBase availability)
            : base(availability)
        { }

        public override IList<Availability> GetAvailability()
        {
            IList<Availability> baseAvailability = base._Availability.GetAvailability();
            IList<Availability> resultAvailability = null;

            if (baseAvailability != null)
            {
                //Sort Availability.
                IEnumerable<Availability> avai = baseAvailability.OrderBy(avail => avail.flight_id)
                                                                .ThenBy(avail => avail.total_adult_fare);

                //Find lowest fare.
                Guid flightId = Guid.Empty;
                Guid transitFlightId = Guid.Empty;
                
                if (avai != null)
                {
                    //Initialize avaiilability result object.
                    resultAvailability = new List<Availability>();

                    //Fill new filter value to list to fill the lowest fare.
                    foreach (Availability a in avai)
                    {
                        if (a.full_flight_flag == 0)
                        {
                            if (flightId != a.flight_id | transitFlightId != a.transit_flight_id)
                            {
                                flightId = a.flight_id;
                                transitFlightId = a.transit_flight_id;

                                //Fill availability of the selected booking class.
                                foreach (Availability b in avai)
                                {
                                    if (b.full_flight_flag == 0)
                                    {
                                        if (a.booking_class_rcd == b.booking_class_rcd &&
                                            a.flight_id == b.flight_id &&
                                            a.transit_flight_id == b.transit_flight_id)
                                        {
                                            resultAvailability.Add(b);
                                        }
                                    }

                                }
                            }
                        }
                    }
                }

                //***************************************************************
                //  Check whether availability result is found.
                //  If yes use availability result.

                if (resultAvailability != null && resultAvailability.Count > 0)
                {
                    return resultAvailability;
                }
                else
                {
                    return baseAvailability;
                }

            }
            else
            {
                return null;
            }
        }
    }
}
