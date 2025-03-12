using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.Serialization;
using System.ServiceModel;

namespace Avantik.Web.Service.Message.Booking
{
    public class Passenger
    {
        #region General Information
 	 	 [MessageBodyMember] 
        public string TitleRcd { get; set; }
 	 	 [MessageBodyMember] 
        public string Lastname { get; set; }
 	 	 [MessageBodyMember] 
        public string Firstname { get; set; }
 	 	 [MessageBodyMember] 
        public string Middlename { get; set; }
 	 	 [MessageBodyMember] 
        public string NationalityRcd { get; set; }
 	 	 [MessageBodyMember] 
        public decimal PassengerWeight { get; set; }
 	 	 [MessageBodyMember] 
        public string GenderTypeRcd { get; set; }
 	 	 [MessageBodyMember] 
        public string PassengerTypeRcd { get; set; }
 	 	 [MessageBodyMember] 
        public Guid ClientProfileId { get; set; }
 	 	 [MessageBodyMember] 
        public long ClientNumber { get; set; }
        #endregion

        #region Address information
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
        public string ZipCode { get; set; }
 	 	 [MessageBodyMember] 
        public string PoBox { get; set; }
 	 	 [MessageBodyMember] 
        public string CountryRcd { get; set; }
 	 	 [MessageBodyMember] 
        public string Street { get; set; }
 	 	 [MessageBodyMember] 
        public string City { get; set; }
        #endregion

        #region Document Information
 	 	 [MessageBodyMember] 
        public string DocumentTypeRcd { get; set; }
 	 	 [MessageBodyMember] 
        public string DocumentNumber { get; set; }
 	 	 [MessageBodyMember] 
        public string ResidenceCountryRcd { get; set; }
 	 	 [MessageBodyMember] 
        public string PassportNumber { get; set; }
         [MessageBodyMember]
        public DateTime PassportIssueDate { get; set; }
         [MessageBodyMember]
        public DateTime PassportExpiryDate { get; set; }
 	 	 [MessageBodyMember] 
        public string PassportIssuePlace { get; set; }
 	 	 [MessageBodyMember] 
        public string PassportBirthPlace { get; set; }
         [MessageBodyMember]
        public DateTime DateOfBirth { get; set; }
 	 	 [MessageBodyMember] 
        public string PassportIssueCountryRcd { get; set; }
         [MessageBodyMember]
         public string CountryCodeLong { get; set; }
        
        #endregion

        #region Contact Information
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
        public string ReceivedFrom { get; set; }
 	 	 [MessageBodyMember] 
        #endregion

        #region Update Information
        public Guid CreateBy { get; set; }
         [MessageBodyMember]
        public DateTime CreateDateTime { get; set; }
 	 	 [MessageBodyMember] 
        public Guid UpdateBy { get; set; }
         [MessageBodyMember]
        public DateTime UpdateDateTime { get; set; }
        #endregion

        #region Properties
         [MessageBodyMember]
         public Guid PassengerId { get; set; }
         [MessageBodyMember]
         public Guid BookingId { get; set; }
         [MessageBodyMember]
         public Guid PassengerProfileId { get; set; }
         [MessageBodyMember]
         public Guid GuardianPassengerId { get; set; }
         [MessageBodyMember]
         public string EmployeeNumber { get; set; }
         [MessageBodyMember]
         public string PassengerRoleRcd { get; set; }
         [MessageBodyMember]
         public string MemberLevelRcd { get; set; }
         [MessageBodyMember]
         public string MemberNumber { get; set; }
         [MessageBodyMember]
         public string RedressNumber { get; set; }
         [MessageBodyMember]
         public string PnrName { get; set; }
         [MessageBodyMember]
         public string MemberAirlineRcd { get; set; }
         [MessageBodyMember]
         public byte WheelchairFlag { get; set; }
         [MessageBodyMember]
         public byte VipFlag { get; set; }
         [MessageBodyMember]
         public byte WindowSeatFlag { get; set; }
        
        
        #endregion

    }
}
