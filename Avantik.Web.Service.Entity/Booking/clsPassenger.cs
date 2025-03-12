using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Avantik.Web.Service.Entity.Booking

{
    public class Passenger 
    {
        #region General Information

        public string TitleRcd { get; set; }
        public string Lastname { get; set; }
        public string Firstname { get; set; }
        public string Middlename { get; set; }
        public string NationalityRcd { get; set; }
        public decimal PassengerWeight { get; set; }
        public string GenderTypeRcd { get; set; }
        public string PassengerTypeRcd { get; set; }
        public Guid ClientProfileId { get; set; }
        public long ClientNumber { get; set; }
        #endregion
        #region Address information
        public string AddressLine1 { get; set; }
        public string AddressLine2 { get; set; }
        public string State { get; set; }
        public string District { get; set; }
        public string Province { get; set; }
        public string ZipCode { get; set; }
        public string PoBox { get; set; }
        public string CountryRcd { get; set; }
        public string Street { get; set; }
        public string City { get; set; }
        #endregion
        #region Document Information
        public string DocumentTypeRcd { get; set; }
        public string DocumentNumber { get; set; }
        public string ResidenceCountryRcd { get; set; }
        public string PassportNumber { get; set; }
        public DateTime PassportIssueDate { get; set; }
        public DateTime PassportExpiryDate { get; set; }
        public string PassportIssuePlace { get; set; }
        public string PassportBirthPlace { get; set; }
        public DateTime DateOfBirth { get; set; }
        public string PassportIssueCountryRcd { get; set; }
        public string CountryCodeLong { get; set; }
        #endregion
        #region Contact Information
        public string ContactName { get; set; }
        public string ContactEmail { get; set; }
        public string MobileEmail { get; set; }
        public string PhoneMobile { get; set; }
        public string PhoneHome { get; set; }
        public string PhoneFax { get; set; }
        public string PhoneBusiness { get; set; }
        public string ReceivedFrom { get; set; }
        #endregion
        #region Update Information
        public Guid CreateBy { get; set; }
        public DateTime CreateDateTime { get; set; }
        public Guid UpdateBy { get; set; }
        public DateTime UpdateDateTime { get; set; }
        #endregion
        #region Properties
        public Guid PassengerId { get; set; }
        public Guid BookingId { get; set; }
        public Guid PassengerProfileId { get; set; }

        public string EmployeeNumber { get; set; }
        public string PassengerRoleRcd { get; set; }
        public string MemberLevelRcd { get; set; }
        public string MemberNumber { get; set; }
        public string RedressNumber { get; set; }
        public string PnrName { get; set; }
        public byte WheelchairFlag { get; set; }
        public byte VipFlag { get; set; }
        public byte WindowSeatFlag { get; set; }
        public Guid GuardianPassengerId { get; set; }
        public string MemberAirlineRcd { get; set; }

        #endregion
    }
}
