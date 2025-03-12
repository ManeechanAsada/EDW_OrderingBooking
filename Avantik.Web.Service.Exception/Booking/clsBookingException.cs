using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace Avantik.Web.Service.Exception.Booking
{
    [Serializable]
    public class BookingException : System.Exception
    {
        public BookingException()
        {
        }

        public BookingException(string message)
            : base(message)
        {
        }

        public BookingException(string message, System.Exception inner)
            : base(message, inner)
        {
        }

    }
}
