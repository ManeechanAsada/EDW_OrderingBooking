using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace Avantik.Web.Service.Exception.Booking
{
    [Serializable]
    public class AgentLogonException : System.Exception
    {
        public AgentLogonException()
        {
        }

        public AgentLogonException(string message)
            : base(message)
        {
        }

        public AgentLogonException(string message, System.Exception inner)
            : base(message, inner)
        {
        }

    }
}
