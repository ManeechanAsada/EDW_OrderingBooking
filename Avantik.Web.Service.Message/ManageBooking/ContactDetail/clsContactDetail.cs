using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.Serialization;
using System.ServiceModel;

namespace Avantik.Web.Service.Message
{
    public class ContactDetail
    {

         [MessageBodyMember] 
         public string ContactName { get; set; }
 	 	 [MessageBodyMember]
         public string ContactEmail { get; set; }
 	 	 [MessageBodyMember] 
         public string PhoneMobile { get; set; }
         [MessageBodyMember]
         public string PhoneHome { get; set; }
         [MessageBodyMember]
         public string PhoneBusiness { get; set; }
         [MessageBodyMember]
         public string PhoneFax { get; set; }
         [MessageBodyMember]
         public string MobileEmail { get; set; }
         [MessageBodyMember]
         public string Language { get; set; }

    }
}
