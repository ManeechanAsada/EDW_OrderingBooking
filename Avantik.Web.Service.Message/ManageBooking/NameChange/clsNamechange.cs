using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.ServiceModel;
using System.Text;
namespace Avantik.Web.Service.Message.ManageBooking
{
    public class NameChange
    {
        [MessageBodyMember]
        public string PassengerId { get; set; }
        [MessageBodyMember]
        public string TitleRcd { get; set; }
        [MessageBodyMember]
        public string Lastname { get; set; }
        [MessageBodyMember]
        public string Firstname { get; set; }
        [MessageBodyMember]
        public string Middlename { get; set; }
        [MessageBodyMember]
        public string GenderTypeRcd { get; set; }
        [MessageBodyMember]
        public DateTime DateOfBirth { get; set; }
    }
}
