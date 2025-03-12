using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.Serialization;
using System.ServiceModel;

namespace Avantik.Web.Service.Message.Client
{
    public class PassengerProfile
    {
        [MessageBodyMember]
        public Guid ClientProfileId { get; set; }
        [MessageBodyMember]
        public Guid PassengerProfileId { get; set; }
        [MessageBodyMember]
        public string PassengerRoleRcd { get; set; }
        [MessageBodyMember]
        public string TitleRcd { get; set; }
        [MessageBodyMember]
        public string Lastname { get; set; }
        [MessageBodyMember]
        public string Firstname { get; set; }
        [MessageBodyMember]
        public string Middlename { get; set; }
        [MessageBodyMember]
        public string LanguageRcd { get; set; }
        [MessageBodyMember]
        public string NationalityRcd { get; set; }
        [MessageBodyMember]
        public string PassengerWeight { get; set; }
        [MessageBodyMember]
        public string GenderTypeRcd { get; set; }
        [MessageBodyMember]
        public string PassengerTypeRcd { get; set; }
        [MessageBodyMember]
        public string DocumentTypeRcd { get; set; }
        [MessageBodyMember]
        public string PassportNumber { get; set; }
        [MessageBodyMember]
        public string PassportIssueDate { get; set; }
        [MessageBodyMember]
        public string PassportExpiryDate { get; set; }
        [MessageBodyMember]
        public string PassportIssuePlace { get; set; }
        [MessageBodyMember]
        public string PassportBirthPlace { get; set; }
        [MessageBodyMember]
        public string DateOfBirth { get; set; }
        [MessageBodyMember]
        public string PassportIssueCountryRcd { get; set; }
        [MessageBodyMember]
        public string ContactName { get; set; }
        [MessageBodyMember]
        public string ContactEmail { get; set; }
        [MessageBodyMember]
        public string MobileEmail { get; set; }
        [MessageBodyMember]
        public string PhoneMobile { get; set; }
        [MessageBodyMember]
        public string PhoneHome { get; set; }
        [MessageBodyMember]
        public string PhoneFax { get; set; }
        [MessageBodyMember]
        public string PhoneBusiness { get; set; }
        [MessageBodyMember]
        public string EmployeeNumber { get; set; }
        [MessageBodyMember]
        public byte WheelchairFlag { get; set; }
        [MessageBodyMember]
        public byte VipFlag { get; set; }
        [MessageBodyMember]
        public string MemberLevelRcd { get; set; }
        [MessageBodyMember]
        public string MemberNumber { get; set; }
        [MessageBodyMember]
        public byte WindowSeatFlag { get; set; }
        [MessageBodyMember]
        public string RedressNumber { get; set; }

    }
}
