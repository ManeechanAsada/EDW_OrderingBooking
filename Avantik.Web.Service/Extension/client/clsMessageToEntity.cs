using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Avantik.Web.Service.Entity.Booking;

namespace Avantik.Web.Service.Extension
{
    public static class MessageToEntity
    {
        public static Entity.Client.Client ToEntityClient(this  Message.Client.Client messageClient, Guid clientProfileId)
        {
            Entity.Client.Client entityClient = null;

            if (messageClient != null)
            {
                entityClient = new Entity.Client.Client();

                entityClient.ClientProfileId = clientProfileId;

                entityClient.StatusCode = messageClient.StatusCode;
                entityClient.ClientNumber = messageClient.ClientNumber;
                entityClient.ClientPassword = messageClient.ClientPassword;
                entityClient.CompanyFlag = messageClient.CompanyFlag;
                entityClient.ProfileOnHoldDateTime = messageClient.ProfileOnHoldDateTime;
                entityClient.ProfileOnHoldComment = messageClient.ProfileOnHoldComment;
                entityClient.ProfileOnHoldBy = messageClient.ProfileOnHoldBy;
                entityClient.CompanyClientProfileId = messageClient.CompanyClientProfileId;
                entityClient.FfpTotal = messageClient.FfpTotal;
                entityClient.FfpPeriod = messageClient.FfpPeriod;
                entityClient.FfpBalance = messageClient.FfpBalance;
                entityClient.ClientTypeRcd = messageClient.ClientTypeRcd;
                entityClient.MemberSinceDate = messageClient.MemberSinceDate;
                entityClient.MemberLevelDisplayName = messageClient.MemberLevelDisplayName;
                entityClient.KeepPoint = messageClient.KeepPoint;

                entityClient.TitleRcd = messageClient.TitleRcd;
                entityClient.Lastname = messageClient.Lastname;
                entityClient.Firstname = messageClient.Firstname;
                entityClient.Middlename = messageClient.Middlename;
                entityClient.LanguageRcd = messageClient.LanguageRcd;
                entityClient.NationalityRcd = messageClient.NationalityRcd;
                entityClient.PassengerWeight = messageClient.PassengerWeight;
                entityClient.GenderTypeRcd = messageClient.GenderTypeRcd;
                entityClient.PassengerTypeRcd = messageClient.PassengerTypeRcd;
                entityClient.AddressLine1 = messageClient.AddressLine1;
                entityClient.AddressLine2 = messageClient.AddressLine2;
                entityClient.State = messageClient.State;
                entityClient.District = messageClient.District;
                entityClient.Province = messageClient.Province;
                entityClient.ZipCode = messageClient.ZipCode;
                entityClient.PoBox = messageClient.PoBox;
                entityClient.CountryRcd = messageClient.CountryRcd;
                entityClient.Street = messageClient.Street;
                entityClient.City = messageClient.City;
                entityClient.DocumentTypeRcd = messageClient.DocumentTypeRcd;
                entityClient.DocumentNumber = messageClient.DocumentNumber;
                entityClient.ResidenceCountryRcd = messageClient.ResidenceCountryRcd;

                entityClient.PassportNumber = messageClient.PassportNumber;
                entityClient.PassportIssueDate = messageClient.PassportIssueDate;
                entityClient.PassportExpiryDate = messageClient.PassportExpiryDate;
                entityClient.PassportIssuePlace = messageClient.PassportIssuePlace;
                entityClient.PassportBirthPlace = messageClient.PassportBirthPlace;
                entityClient.DateOfBirth = messageClient.DateOfBirth;
                entityClient.PassportIssueCountryRcd = messageClient.PassportIssueCountryRcd;
                entityClient.ContactName = messageClient.ContactName;
                entityClient.ContactEmail = messageClient.ContactEmail;
                entityClient.MobileEmail = messageClient.MobileEmail;
                entityClient.PhoneMobile = messageClient.PhoneMobile;
                entityClient.PhoneHome = messageClient.PhoneHome;
                entityClient.PhoneFax = messageClient.PhoneFax;
                entityClient.PhoneBusiness = messageClient.PhoneBusiness;
                entityClient.EmployeeNumber = messageClient.EmployeeNumber;
                entityClient.WheelchairFlag = messageClient.WheelchairFlag;
                entityClient.VipFlag = messageClient.VipFlag;
                entityClient.MemberLevelRcd = messageClient.MemberLevelRcd;
                entityClient.MemberNumber = messageClient.MemberNumber;
                entityClient.WindowSeatFlag = messageClient.WindowSeatFlag;
                entityClient.RedressNumber = messageClient.RedressNumber;

            }

            return entityClient;
        }

        public static IList<Entity.Client.PassengerProfile> ToListEntityClient(this  IList<Message.Client.PassengerProfile> p,Guid clientProfileId)
        {
            IList<Entity.Client.PassengerProfile> passengerList = null;

            if (p != null)
            {
                passengerList = new List<Entity.Client.PassengerProfile>();
                for (int i = 0; i < p.Count; i++)
                {
                    passengerList.Add(p[i].ToEntityClient(clientProfileId));
                }
            }

            return passengerList;
        }
        public static Entity.Client.PassengerProfile ToEntityClient(this  Message.Client.PassengerProfile p, Guid clientProfileId)
        {
            Entity.Client.PassengerProfile passengerProfile = null;
            if (p != null)
            {
                passengerProfile = new Entity.Client.PassengerProfile();

                passengerProfile.ClientProfileId = clientProfileId;

                if (p.PassengerProfileId != Guid.Empty)
                    passengerProfile.PassengerProfileId = p.PassengerProfileId;
                else
                    passengerProfile.PassengerProfileId = Guid.NewGuid();

                passengerProfile.PassengerRoleRcd = p.PassengerRoleRcd;
                passengerProfile.TitleRcd = p.TitleRcd;
                passengerProfile.Lastname = p.Lastname;
                passengerProfile.Firstname = p.Firstname;
                passengerProfile.Middlename = p.Middlename;
                passengerProfile.LanguageRcd = p.LanguageRcd;
                passengerProfile.NationalityRcd = p.NationalityRcd;
                passengerProfile.PassengerWeight = p.PassengerWeight;
                passengerProfile.GenderTypeRcd = p.GenderTypeRcd;
                passengerProfile.PassengerTypeRcd = p.PassengerTypeRcd;
                passengerProfile.DocumentTypeRcd = p.DocumentTypeRcd;
                passengerProfile.PassportNumber = p.PassportNumber;
                passengerProfile.PassportIssueDate = p.PassportIssueDate;
                passengerProfile.PassportExpiryDate = p.PassportExpiryDate;
                passengerProfile.PassportIssuePlace = p.PassportIssuePlace;
                passengerProfile.PassportBirthPlace = p.PassportBirthPlace;
                passengerProfile.DateOfBirth = p.DateOfBirth;
                passengerProfile.PassportIssueCountryRcd = p.PassportIssueCountryRcd;
                passengerProfile.ContactName = p.ContactName;
                passengerProfile.ContactEmail = p.ContactEmail;
                passengerProfile.MobileEmail = p.MobileEmail;
                passengerProfile.PhoneMobile = p.PhoneMobile;
                passengerProfile.PhoneHome = p.PhoneHome;
                passengerProfile.PhoneFax = p.PhoneFax;
                passengerProfile.PhoneBusiness = p.PhoneBusiness;
                passengerProfile.EmployeeNumber = p.EmployeeNumber;
                passengerProfile.WheelchairFlag = p.WheelchairFlag;
                passengerProfile.VipFlag = p.VipFlag;
                passengerProfile.MemberLevelRcd = p.MemberLevelRcd;
                passengerProfile.MemberNumber = p.MemberNumber;
                passengerProfile.WindowSeatFlag = p.WindowSeatFlag;
                passengerProfile.RedressNumber = p.RedressNumber;
            }

            return passengerProfile;
        }

    }
}
