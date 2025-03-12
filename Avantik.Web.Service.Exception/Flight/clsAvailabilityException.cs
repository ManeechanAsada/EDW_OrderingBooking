using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace Avantik.Web.Service.Exception.Flight
{
    [Serializable]
    public class AvailabilityException : System.Exception
    {
        public AvailabilityException()
        {
        }

        public AvailabilityException(string message)
            : base(message)
        {
        }

        public AvailabilityException(string message, System.Exception inner)
            : base(message, inner)
        {
        }
    }
}
