using Avantik.Web.Service.COMHelper;
using Avantik.Web.Service.Entity.Booking;
using Avantik.Web.Service.Entity.Client;
using Avantik.Web.Service.Exception.Booking;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Avantik.Web.Service.Model.COM.Extension
{
    public static class clsClientRecordsetObject
    {
        public static Client FillClientObject(this Client client, ADODB.Recordset rs)
        {
            client = new Client();
            try
            {
                client.ClientProfileId = RecordsetHelper.ToGuid(rs, "client_profile_id");
                client.StatusCode = RecordsetHelper.ToString(rs, "status_code");
                client.ClientNumber = RecordsetHelper.ToString(rs, "client_number");
                client.ClientPassword = RecordsetHelper.ToString(rs, "client_password");
                client.CompanyFlag = RecordsetHelper.ToBoolean(rs, "company_flag");
                client.ProfileOnHoldDateTime = RecordsetHelper.ToDateTime(rs, "profile_on_hold_date_time");
                client.ProfileOnHoldComment = RecordsetHelper.ToString(rs, "profile_on_hold_comment");
                client.ProfileOnHoldBy = RecordsetHelper.ToGuid(rs, "profile_on_hold_by");
                client.CompanyClientProfileId = RecordsetHelper.ToGuid(rs, "company_client_profile_id");
                client.FfpTotal = RecordsetHelper.ToDouble(rs, "ffp_total");
                client.FfpPeriod = RecordsetHelper.ToDouble(rs, "ffp_period");
                client.FfpBalance = RecordsetHelper.ToDouble(rs, "ffp_balance");
                client.ClientTypeRcd = RecordsetHelper.ToString(rs, "client_type_rcd");
                client.MemberSinceDate = RecordsetHelper.ToDateTime(rs, "member_since_date");
                client.MemberLevelDisplayName = RecordsetHelper.ToString(rs, "member_level_display_name");
                client.KeepPoint = RecordsetHelper.ToDouble(rs, "keep_point");

                client.TitleRcd = RecordsetHelper.ToString(rs, "title_rcd");
                client.Lastname = RecordsetHelper.ToString(rs, "Lastname");
                client.Firstname = RecordsetHelper.ToString(rs, "Firstname");
                client.Middlename = RecordsetHelper.ToString(rs, "Middlename");
                client.LanguageRcd = RecordsetHelper.ToString(rs, "language_rcd");
                client.NationalityRcd = RecordsetHelper.ToString(rs, "nationality_rcd");
                client.PassengerWeight = RecordsetHelper.ToString(rs, "passenger_weight");
                client.GenderTypeRcd = RecordsetHelper.ToString(rs, "gender_type_rcd");
                client.PassengerTypeRcd = RecordsetHelper.ToString(rs, "passenger_type_rcd");
                client.AddressLine1 = RecordsetHelper.ToString(rs, "address_line1");
                client.AddressLine2 = RecordsetHelper.ToString(rs, "address_line2");
                client.State = RecordsetHelper.ToString(rs, "State");
                client.District = RecordsetHelper.ToString(rs, "District");
                client.Province = RecordsetHelper.ToString(rs, "Province");
                client.ZipCode = RecordsetHelper.ToBoolean(rs, "zip_code");
                client.PoBox = RecordsetHelper.ToString(rs, "po_box");
                client.CountryRcd = RecordsetHelper.ToBoolean(rs, "country_rcd");
                client.Street = RecordsetHelper.ToString(rs, "Street");
                client.City = RecordsetHelper.ToString(rs, "City");
                client.DocumentTypeRcd = RecordsetHelper.ToString(rs, "document_type_rcd");
                client.DocumentNumber = RecordsetHelper.ToString(rs, "document_number");
                client.ResidenceCountryRcd = RecordsetHelper.ToString(rs, "residence_country_rcd");

                client.PassportNumber = RecordsetHelper.ToString(rs, "passport_number");
                client.PassportIssueDate = RecordsetHelper.ToString(rs, "passport_issue_date");
                client.PassportExpiryDate = RecordsetHelper.ToString(rs, "passport_expiry_date");
                client.PassportIssuePlace = RecordsetHelper.ToString(rs, "passport_issue_place");
                client.PassportBirthPlace = RecordsetHelper.ToString(rs, "passport_birth_place");
                client.DateOfBirth = RecordsetHelper.ToString(rs, "date_of_birth");
                client.PassportIssueCountryRcd = RecordsetHelper.ToString(rs, "passport_issue_country_rcd");
                client.ContactName = RecordsetHelper.ToString(rs, "contact_name");
                client.ContactEmail = RecordsetHelper.ToString(rs, "contact_email");
                client.MobileEmail = RecordsetHelper.ToString(rs, "mobile_email");
                client.PhoneMobile = RecordsetHelper.ToString(rs, "phone_mobile");
                client.PhoneHome = RecordsetHelper.ToString(rs, "phone_home");
                client.PhoneFax = RecordsetHelper.ToString(rs, "phone_fax");
                client.PhoneBusiness = RecordsetHelper.ToString(rs, "phone_business");
                client.EmployeeNumber = RecordsetHelper.ToString(rs, "employee_number");
                client.WheelchairFlag = RecordsetHelper.ToByte(rs, "wheelchair_flag");
                client.VipFlag = RecordsetHelper.ToByte(rs, "vip_flag");
                client.MemberLevelRcd = RecordsetHelper.ToString(rs, "member_level_rcd");
                client.MemberNumber = RecordsetHelper.ToString(rs, "member_number");
                client.WindowSeatFlag = RecordsetHelper.ToByte(rs, "window_seat_flag");
                client.RedressNumber = RecordsetHelper.ToString(rs, "redress_number");

            }
            catch
            {
                throw;
            }
            return client;
        }

        public static IList<PassengerProfile> FillClientObject(this IList<PassengerProfile> passengerList, ADODB.Recordset rs)
        {
            if (rs != null && rs.RecordCount > 0)
            {
                passengerList = new List<PassengerProfile>();
                PassengerProfile passengerProfile = null;

                try
                {
                    rs.MoveFirst();
                    while (!rs.EOF)
                    {
                        passengerProfile = new PassengerProfile();

                        passengerProfile.ClientProfileId = RecordsetHelper.ToGuid(rs, "client_profile_id");
                        passengerProfile.PassengerProfileId = RecordsetHelper.ToGuid(rs, "passenger_profile_id");
                        passengerProfile.PassengerRoleRcd = RecordsetHelper.ToString(rs, "passenger_role_rcd");
                        passengerProfile.TitleRcd = RecordsetHelper.ToString(rs, "title_rcd");
                        passengerProfile.Lastname = RecordsetHelper.ToString(rs, "Lastname");
                        passengerProfile.Firstname = RecordsetHelper.ToString(rs, "Firstname");
                        passengerProfile.Middlename = RecordsetHelper.ToString(rs, "Middlename");
                        passengerProfile.LanguageRcd = RecordsetHelper.ToString(rs, "language_rcd");
                        passengerProfile.NationalityRcd = RecordsetHelper.ToString(rs, "nationality_rcd");
                        passengerProfile.PassengerWeight = RecordsetHelper.ToString(rs, "passenger_weight");
                        passengerProfile.GenderTypeRcd = RecordsetHelper.ToString(rs, "gender_type_rcd");
                        passengerProfile.PassengerTypeRcd = RecordsetHelper.ToString(rs, "passenger_type_rcd");
                        passengerProfile.DocumentTypeRcd = RecordsetHelper.ToString(rs, "document_type_rcd");
                        passengerProfile.PassportNumber = RecordsetHelper.ToString(rs, "passport_number");
                        passengerProfile.PassportIssueDate = RecordsetHelper.ToString(rs, "passport_issue_date");
                        passengerProfile.PassportExpiryDate = RecordsetHelper.ToString(rs, "passport_expiry_date");
                        passengerProfile.PassportIssuePlace = RecordsetHelper.ToString(rs, "passport_issue_place");
                        passengerProfile.PassportBirthPlace = RecordsetHelper.ToString(rs, "passport_birth_place");
                        passengerProfile.DateOfBirth = RecordsetHelper.ToString(rs, "date_of_birth");
                        passengerProfile.PassportIssueCountryRcd = RecordsetHelper.ToString(rs, "passport_issue_country_rcd");
                        passengerProfile.ContactName = RecordsetHelper.ToString(rs, "contact_name");
                        passengerProfile.ContactEmail = RecordsetHelper.ToString(rs, "contact_email");
                        passengerProfile.MobileEmail = RecordsetHelper.ToString(rs, "mobile_email");
                        passengerProfile.PhoneMobile = RecordsetHelper.ToString(rs, "phone_mobile");
                        passengerProfile.PhoneHome = RecordsetHelper.ToString(rs, "phone_home");
                        passengerProfile.PhoneFax = RecordsetHelper.ToString(rs, "phone_fax");
                        passengerProfile.PhoneBusiness = RecordsetHelper.ToString(rs, "phone_business");
                        passengerProfile.EmployeeNumber = RecordsetHelper.ToString(rs, "employee_number");
                        passengerProfile.WheelchairFlag = RecordsetHelper.ToByte(rs, "wheelchair_flag");
                        passengerProfile.VipFlag = RecordsetHelper.ToByte(rs, "vip_flag");
                        passengerProfile.MemberLevelRcd = RecordsetHelper.ToString(rs, "member_level_rcd");
                        passengerProfile.MemberNumber = RecordsetHelper.ToString(rs, "member_number");
                        passengerProfile.WindowSeatFlag = RecordsetHelper.ToByte(rs, "window_seat_flag");
                        passengerProfile.RedressNumber = RecordsetHelper.ToString(rs, "redress_number");

                        passengerList.Add(passengerProfile);

                        rs.MoveNext();
                    }

                }
                catch
                {
                    throw;
                }

            }
            return passengerList;

        }

    }
}
