using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.Serialization;
using System.ServiceModel;

namespace Avantik.Web.Service.Message.Booking
{
    public class Phone 
    {
        #region Property
        [MessageBodyMember]
        public string PhoneMobile { get; set; }
        [MessageBodyMember]
        public string PhoneHome { get; set; }
        [MessageBodyMember]
        public string PhoneBusiness { get; set; }
        [MessageBodyMember]
        public string PhoneFax { get; set; }

        #endregion
    }
}
