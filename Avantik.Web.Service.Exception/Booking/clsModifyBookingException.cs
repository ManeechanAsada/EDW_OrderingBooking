using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace Avantik.Web.Service.Exception.Booking
{
    [Serializable]
    public class ModifyBookingException : System.Exception
    {
        string _ErrorCode = string.Empty;
        public string ErrorCode
        {
            get { return _ErrorCode; }
        }
        public ModifyBookingException()
        {
        }

        public ModifyBookingException(string message)
            : base(message)
        {
        }
        public ModifyBookingException(string message,string errorCode)
            : base(message)
        {
            _ErrorCode = errorCode;
        }
        public ModifyBookingException(string message, System.Exception inner)
            : base(message, inner)
        {
        }
    }
}
