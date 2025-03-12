using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.Serialization;
using System.ServiceModel;

namespace Avantik.Web.Service.Message.Client
{
    public class Client 
    {
        [MessageBodyMember]
        public Guid ClientProfileId { get; set; }
        [MessageBodyMember]
        public string StatusCode { get; set; }
        [MessageBodyMember]
        public string ClientNumber { get; set; }
        [MessageBodyMember]
        public string ClientPassword { get; set; }
        [MessageBodyMember]
        public bool CompanyFlag { get; set; }
        [DataMember(EmitDefaultValue = false)]
        public DateTime ProfileOnHoldDateTime { get; set; }
        [MessageBodyMember]
        public string ProfileOnHoldComment { get; set; }
        [MessageBodyMember]
        public Guid ProfileOnHoldBy { get; set; }
        [MessageBodyMember]
        public Guid CompanyClientProfileId { get; set; }
        [MessageBodyMember]
        public double FfpTotal { get; set; }
        [MessageBodyMember]
        public double FfpPeriod { get; set; }
        [MessageBodyMember]
        public double FfpBalance { get; set; }
        [MessageBodyMember]
        public string ClientTypeRcd { get; set; }
        [DataMember(EmitDefaultValue = false)]
        public DateTime MemberSinceDate { get; set; }
        [MessageBodyMember]
        public string MemberLevelDisplayName { get; set; }
        [MessageBodyMember]
        public double KeepPoint { get; set; }

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
        public string AddressLine1 { get; set; }
        [MessageBodyMember]
        public string AddressLine2 { get; set; }
        [MessageBodyMember]
        public string State { get; set; }
        [MessageBodyMember]
        public string District { get; set; }
        [MessageBodyMember]
        public string Province { get; set; }
        [MessageBodyMember]
        public bool ZipCode { get; set; }
        [MessageBodyMember]
        public string PoBox { get; set; }
        [MessageBodyMember]
        public bool CountryRcd { get; set; }
        [MessageBodyMember]
        public string Street { get; set; }
        [MessageBodyMember]
        public string City { get; set; }
        [MessageBodyMember]
        public string DocumentTypeRcd { get; set; }
        [MessageBodyMember]
        public string DocumentNumber { get; set; }
        [MessageBodyMember]
        public string ResidenceCountryRcd { get; set; }
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
