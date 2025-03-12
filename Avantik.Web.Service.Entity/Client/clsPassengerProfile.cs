using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

namespace Avantik.Web.Service.Entity.Client
{
    public class PassengerProfile
    {
        public Guid ClientProfileId { get; set; }
        public Guid PassengerProfileId { get; set; }
        public string PassengerRoleRcd { get; set; }
        public string TitleRcd { get; set; }
        public string Lastname { get; set; }
        public string Firstname { get; set; }
        public string Middlename { get; set; }
        public string LanguageRcd { get; set; }
        public string NationalityRcd { get; set; }
        public string PassengerWeight { get; set; }
        public string GenderTypeRcd { get; set; }
        public string PassengerTypeRcd { get; set; }
        public string DocumentTypeRcd { get; set; }
        public string PassportNumber { get; set; }
        public string PassportIssueDate { get; set; }
        public string PassportExpiryDate { get; set; }
        public string PassportIssuePlace { get; set; }
        public string PassportBirthPlace { get; set; }
        public string DateOfBirth { get; set; }
        public string PassportIssueCountryRcd { get; set; }
        public string ContactName { get; set; }
        public string ContactEmail { get; set; }
        public string MobileEmail { get; set; }
        public string PhoneMobile { get; set; }
        public string PhoneHome { get; set; }
        public string PhoneFax { get; set; }
        public string PhoneBusiness { get; set; }
        public string EmployeeNumber { get; set; }
        public byte WheelchairFlag { get; set; }
        public byte VipFlag { get; set; }
        public string MemberLevelRcd { get; set; }
        public string MemberNumber { get; set; }
        public byte WindowSeatFlag { get; set; }
        public string RedressNumber { get; set; }

    }
}
