using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.Serialization;
using System.ServiceModel;

namespace Avantik.Web.Service.Message
{
    public class PassengerInfo
    {
        [MessageBodyMember]
        public string PassengerId { get; set; }
        [MessageBodyMember] 
        public String DocumentTypeRcd { get; set; }
 	 	 [MessageBodyMember] 
        public string DocumentNumber { get; set; }
 	 	 [MessageBodyMember] 
        public string PassportBirthPlace { get; set; }
         [MessageBodyMember]
        public DateTime PassportIssueDate { get; set; }
         [MessageBodyMember]
        public DateTime PassportExpiryDate { get; set; }
 	 	 [MessageBodyMember] 
        public string PassportIssuePlace { get; set; }
 	 	 [MessageBodyMember] 
        public string PassportIssueCountryRcd { get; set; }
         [MessageBodyMember]
         public string NationalityRcd { get; set; }
    }
}
