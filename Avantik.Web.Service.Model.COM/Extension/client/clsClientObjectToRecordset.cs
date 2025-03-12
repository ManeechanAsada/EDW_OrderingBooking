using Avantik.Web.Service.Entity.Booking;
using Avantik.Web.Service.Entity.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Avantik.Web.Service.COMHelper;

namespace Avantik.Web.Service.Model.COM.Extension
{
    public static class ClientObjectToRecordset
    {
        public static void ToRecordsetClient(this  Client client, ref ADODB.Recordset rs)
        {
            if (client != null )
            {
                try
                {
                  rs.AddNew();

                  rs.Fields["client_profile_id"].Value = client.ClientProfileId.ToRsString();
                  rs.Fields["status_code"].Value = client.StatusCode;
                //  rs.Fields["client_number"].Value = client.ClientNumber;
                  rs.Fields["client_password"].Value = client.ClientPassword;
                  rs.Fields["company_flag"].Value = client.CompanyFlag;
                  rs.Fields["profile_on_hold_date_time"].Value = client.ProfileOnHoldDateTime;
                  rs.Fields["profile_on_hold_comment"].Value = client.ProfileOnHoldComment;
                //  rs.Fields["profile_on_hold_by"].Value = client.ProfileOnHoldBy;
                  rs.Fields["company_client_profile_id"].Value = client.CompanyClientProfileId.ToRsString();
                  rs.Fields["ffp_total"].Value = client.FfpTotal;
                  rs.Fields["ffp_period"].Value = client.FfpPeriod;
                  rs.Fields["ffp_balance"].Value = client.FfpBalance;
                  rs.Fields["client_type_rcd"].Value = client.ClientTypeRcd;
                  rs.Fields["member_since_date"].Value = client.MemberSinceDate;
                //  rs.Fields["member_level_display_name"].Value = client.MemberLevelDisplayName;
                //  rs.Fields["keep_point"].Value = client.KeepPoint;

                  rs.Fields["title_rcd"].Value = client.TitleRcd;
                  rs.Fields["Lastname"].Value = client.Lastname;
                  rs.Fields["Firstname"].Value = client.Firstname;
                  rs.Fields["Middlename"].Value = client.Middlename;
                  rs.Fields["language_rcd"].Value = client.LanguageRcd;
                  rs.Fields["nationality_rcd"].Value = client.NationalityRcd;
                //  rs.Fields["passenger_weight"].Value = client.PassengerWeight;
                  rs.Fields["gender_type_rcd"].Value = client.GenderTypeRcd;
                  rs.Fields["passenger_type_rcd"].Value = client.PassengerTypeRcd;
                  rs.Fields["address_line1"].Value = client.AddressLine1;
                  rs.Fields["address_line2"].Value = client.AddressLine2;
                  rs.Fields["State"].Value = client.State;
                  rs.Fields["District"].Value = client.District;
                  rs.Fields["Province"].Value = client.Province;
                  rs.Fields["zip_code"].Value = client.ZipCode;
                  rs.Fields["po_box"].Value = client.PoBox;
                  rs.Fields["country_rcd"].Value = client.CountryRcd;
                  rs.Fields["Street"].Value = client.Street;
                  rs.Fields["City"].Value = client.City;
                 // rs.Fields["document_type_rcd"].Value = client.DocumentTypeRcd;
                 // rs.Fields["document_number"].Value = client.DocumentNumber;
                 // rs.Fields["residence_country_rcd"].Value = client.ResidenceCountryRcd;

                  rs.Fields["passport_number"].Value = client.PassportNumber;
                  rs.Fields["passport_issue_date"].Value = client.PassportIssueDate;
                  rs.Fields["passport_expiry_date"].Value = client.PassportExpiryDate;
                  rs.Fields["passport_issue_place"].Value = client.PassportIssuePlace;
                  rs.Fields["passport_birth_place"].Value = client.PassportBirthPlace;
                  rs.Fields["date_of_birth"].Value = client.DateOfBirth;
                //  rs.Fields["passport_issue_country_rcd"].Value = client.PassportIssueCountryRcd;
                  rs.Fields["contact_name"].Value = client.ContactName;
                  rs.Fields["contact_email"].Value = client.ContactEmail;
                  rs.Fields["mobile_email"].Value = client.MobileEmail;
                  rs.Fields["phone_mobile"].Value = client.PhoneMobile;
                  rs.Fields["phone_home"].Value = client.PhoneHome;
                  rs.Fields["phone_fax"].Value = client.PhoneFax;
                  rs.Fields["phone_business"].Value = client.PhoneBusiness;
                  rs.Fields["employee_number"].Value = client.EmployeeNumber;
                  rs.Fields["wheelchair_flag"].Value = client.WheelchairFlag;
                  rs.Fields["vip_flag"].Value = client.VipFlag;
                  rs.Fields["member_level_rcd"].Value = client.MemberLevelRcd;
                  rs.Fields["member_number"].Value = client.MemberNumber;
                  rs.Fields["window_seat_flag"].Value = client.WindowSeatFlag;
                 // rs.Fields["redress_number"].Value = client.RedressNumber;


                }
                catch (SystemException ex)
                {
                    throw ex;
                }

            }
        }

        public static void ToRecordsetEditPassenger(this  IList<PassengerProfile> passengers, ref ADODB.Recordset rsPassenger)
        {
            if (passengers != null && passengers.Count > 0)
            {
                try
                {
                    rsPassenger.MoveFirst();
                    while (!rsPassenger.EOF)
                    {
                        if (rsPassenger.Fields["passenger_role_rcd"].Value == "MYSELF")
                        {
                            foreach (PassengerProfile p in passengers)
                            {
                                if (p.PassengerRoleRcd == "MYSELF")
                                {
                                    //Duplicate myself client profile to passenger
                                    rsPassenger.Fields["client_profile_id"].Value = p.ClientProfileId.ToRsString();
                                    rsPassenger.Fields["title_rcd"].Value = p.TitleRcd;
                                    rsPassenger.Fields["Lastname"].Value = p.Lastname;
                                    rsPassenger.Fields["Firstname"].Value = p.Firstname;
                                    rsPassenger.Fields["Middlename"].Value = p.Middlename;
                                    rsPassenger.Fields["nationality_rcd"].Value = p.NationalityRcd;
                                    rsPassenger.Fields["passenger_weight"].Value = p.PassengerWeight;
                                    rsPassenger.Fields["gender_type_rcd"].Value = p.GenderTypeRcd;
                                    rsPassenger.Fields["document_type_rcd"].Value = p.DocumentTypeRcd;
                                    rsPassenger.Fields["passenger_type_rcd"].Value = p.PassengerTypeRcd;
                                    rsPassenger.Fields["passport_number"].Value = p.PassportNumber;
                                    rsPassenger.Fields["passport_issue_date"].Value = p.PassportIssueDate;
                                    rsPassenger.Fields["passport_expiry_date"].Value = p.PassportExpiryDate;
                                    rsPassenger.Fields["passport_issue_place"].Value = p.PassportIssuePlace;
                                    rsPassenger.Fields["passport_birth_place"].Value = p.PassportBirthPlace;
                                    rsPassenger.Fields["date_of_birth"].Value = p.DateOfBirth;
                                    rsPassenger.Fields["passport_issue_country_rcd"].Value = p.PassportIssueCountryRcd;
                                    rsPassenger.Fields["wheelchair_flag"].Value = p.WheelchairFlag;
                                    rsPassenger.Fields["vip_flag"].Value = p.VipFlag;
                                    rsPassenger.Fields["passenger_role_rcd"].Value = p.PassengerRoleRcd;
                                    rsPassenger.Fields["member_level_rcd"].Value = p.MemberLevelRcd;
                                    rsPassenger.Fields["member_number"].Value = p.MemberNumber;
                                    rsPassenger.Fields["window_seat_flag"].Value = p.WindowSeatFlag;
                                    rsPassenger.Fields["redress_number"].Value = p.RedressNumber;
                                }
                            }
                        }

                        rsPassenger.MoveNext();
                    }
                }
                catch
                {
                    throw;
                }
            }
        }

        public static void ToRecordsetClient(this  IList<PassengerProfile> passengers, ref ADODB.Recordset rs)
        {
            if (passengers != null && passengers.Count > 0)
            {
                try
                {
                    foreach (PassengerProfile p in passengers)
                    {
                        //Duplicate myself client profile to passenger
                        rs.AddNew();
                        rs.Fields["client_profile_id"].Value = p.ClientProfileId.ToRsString();

                        rs.Fields["passenger_profile_id"].Value = p.PassengerProfileId.ToRsString();

                        rs.Fields["passenger_role_rcd"].Value = p.PassengerRoleRcd;
                        rs.Fields["title_rcd"].Value = p.TitleRcd;
                        rs.Fields["Lastname"].Value = p.Lastname;
                        rs.Fields["Firstname"].Value = p.Firstname;
                        rs.Fields["Middlename"].Value = p.Middlename;
                       // rs.Fields["language_rcd"].Value = p.LanguageRcd;
                        rs.Fields["nationality_rcd"].Value = p.NationalityRcd;
                        rs.Fields["passenger_weight"].Value = p.PassengerWeight;
                        rs.Fields["gender_type_rcd"].Value = p.GenderTypeRcd;
                        rs.Fields["passenger_type_rcd"].Value = p.PassengerTypeRcd;
                        rs.Fields["document_type_rcd"].Value = p.DocumentTypeRcd;
                        rs.Fields["passport_number"].Value = p.PassportNumber;
                        rs.Fields["passport_issue_date"].Value = p.PassportIssueDate;
                        rs.Fields["passport_expiry_date"].Value = p.PassportExpiryDate;
                        rs.Fields["passport_issue_place"].Value = p.PassportIssuePlace;
                        rs.Fields["passport_birth_place"].Value = p.PassportBirthPlace;
                        rs.Fields["date_of_birth"].Value = p.DateOfBirth;
                        rs.Fields["passport_issue_country_rcd"].Value = p.PassportIssueCountryRcd;
                       // rs.Fields["contact_name"].Value = p.ContactName;
                       // rs.Fields["contact_email"].Value = p.ContactEmail;
                       // rs.Fields["mobile_email"].Value = p.MobileEmail;
                       // rs.Fields["phone_mobile"].Value = p.PhoneMobile;
                       // rs.Fields["phone_home"].Value = p.PhoneHome;
                      //  rs.Fields["phone_fax"].Value = p.PhoneFax;
                      //  rs.Fields["phone_business"].Value = p.PhoneBusiness;
                        rs.Fields["employee_number"].Value = p.EmployeeNumber;
                        rs.Fields["wheelchair_flag"].Value = p.WheelchairFlag;
                        rs.Fields["vip_flag"].Value = p.VipFlag;
                       // rs.Fields["member_level_rcd"].Value = p.MemberLevelRcd;
                       // rs.Fields["member_number"].Value = p.MemberNumber;
                        rs.Fields["window_seat_flag"].Value = p.WindowSeatFlag;
                        rs.Fields["redress_number"].Value = p.RedressNumber;



                    }
                }
                catch
                {
                    throw;
                }
            }
        }

    }
}
