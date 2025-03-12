using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Avantik.Web.Service.Entity.Agency;
using Avantik.Web.Service.COMHelper;
using Avantik.Web.Service.Exception.Booking;

namespace Avantik.Web.Service.Model.COM.Extension
{
    public static class RecordsetObjectAgency
    {
        public static Agent FillAgency(this Agent agent, ADODB.Recordset rs)
        {
            agent = new Agent();
            try
            {
                agent.AgencyCode = RecordsetHelper.ToString(rs, "agency_code");
                agent.CurrencyRcd = RecordsetHelper.ToString(rs, "currency_rcd");
                agent.AgencyPaymentTypeRcd = RecordsetHelper.ToString(rs, "agency_payment_type_rcd");
                agent.AirportRcd = RecordsetHelper.ToString(rs, "airport_rcd");
                agent.CountryRcd = RecordsetHelper.ToString(rs, "country_rcd");
                agent.LanguageRcd = RecordsetHelper.ToString(rs, "language_rcd");
                agent.AgencyPassword = RecordsetHelper.ToString(rs, "agency_password");
                agent.DefaultUserAccountId = RecordsetHelper.ToGuid(rs, "default_user_account_id");
                agent.UserLogon = RecordsetHelper.ToString(rs, "user_logon");
                agent.AgencyLogon = RecordsetHelper.ToString(rs, "agency_logon");
                agent.AgencyName = RecordsetHelper.ToString(rs, "agency_name");
                agent.AgLanguageRcd = RecordsetHelper.ToString(rs, "ag_language_rcd");
                agent.DefaultETicketFlag = RecordsetHelper.ToByte(rs, "default_e_ticket_flag");
                agent.Email = RecordsetHelper.ToString(rs, "email");
                agent.StatusCode = RecordsetHelper.ToString(rs, "status_code");
                agent.MerchantId = RecordsetHelper.ToGuid(rs, "merchant_id");
                agent.NotifyBy = RecordsetHelper.ToString(rs, "notify_by");
                agent.DefaultCustomerDocumentId = RecordsetHelper.ToGuid(rs, "default_customer_document_id");
                agent.DefaultSmallItineraryDocumentId = RecordsetHelper.ToGuid(rs, "default_small_itinerary_document_id");
                agent.DefaultInternalItineraryDocumentId = RecordsetHelper.ToGuid(rs, "default_internal_itinerary_document_id");
                agent.PaymentDefaultCode = RecordsetHelper.ToString(rs, "payment_default_code");
                agent.AgencyTypeCode = RecordsetHelper.ToString(rs, "agency_type_code");
                agent.UserAccountId = RecordsetHelper.ToGuid(rs, "user_account_id");
                agent.UserCode = RecordsetHelper.ToString(rs, "user_code");
                agent.Lastname = RecordsetHelper.ToString(rs, "lastname");
                agent.Middlename = RecordsetHelper.ToString(rs, "middlename");
                agent.Firstname = RecordsetHelper.ToString(rs, "firstname");
                agent.OriginRcd = RecordsetHelper.ToString(rs, "origin_rcd");
                agent.OutstandingInvoice = RecordsetHelper.ToDecimal(rs, "outstanding_invoice");
                agent.BookingPayment = RecordsetHelper.ToDecimal(rs, "booking_payment");
                agent.AgencyAccount = RecordsetHelper.ToDecimal(rs, "agency_account");
                agent.WebsiteAddress = RecordsetHelper.ToString(rs, "website_address");
                agent.TtyAddress = RecordsetHelper.ToString(rs, "tty_address");
                agent.CreateDateTime = RecordsetHelper.ToDateTime(rs, "create_date_time");
                agent.UpdateDateTime = RecordsetHelper.ToDateTime(rs, "update_date_time");
                agent.CashbookClosingRcd = RecordsetHelper.ToString(rs, "cashbook_closing_rcd");
                agent.CashbookClosingId = RecordsetHelper.ToGuid(rs, "cashbook_closing_id");
                agent.CreateBy = RecordsetHelper.ToGuid(rs, "create_by");
                agent.AgencyTimezone = RecordsetHelper.ToInt16(rs, "agency_timezone");
                agent.SystemSettingTimezone = RecordsetHelper.ToInt16(rs, "system_setting_timezone");
                agent.CompanyClientProfileId = RecordsetHelper.ToGuid(rs, "company_client_profile_id");
                agent.ClientProfileId = RecordsetHelper.ToGuid(rs, "client_profile_id");
                agent.InvoiceDays = RecordsetHelper.ToString(rs, "invoice_days");
                agent.AddressLine1 = RecordsetHelper.ToString(rs, "address_line1");
                agent.AddressLine2 = RecordsetHelper.ToString(rs, "address_line2");
                agent.City = RecordsetHelper.ToString(rs, "city");
                agent.BankCode = RecordsetHelper.ToString(rs, "bank_code");
                agent.BankName = RecordsetHelper.ToString(rs, "bank_name");
                agent.BankAccount = RecordsetHelper.ToString(rs, "bank_account");
                agent.ContactPerson = RecordsetHelper.ToString(rs, "contact_person");
                agent.District = RecordsetHelper.ToString(rs, "district");
                agent.Phone = RecordsetHelper.ToString(rs, "phone");
                agent.Fax = RecordsetHelper.ToString(rs, "fax");
                agent.PoBox = RecordsetHelper.ToString(rs, "po_box");
                agent.Province = RecordsetHelper.ToString(rs, "province");
                agent.State = RecordsetHelper.ToString(rs, "state");
                agent.Street = RecordsetHelper.ToString(rs, "street");
                agent.ZipCode = RecordsetHelper.ToString(rs, "zip_code");
                agent.B2bBspFromDate = RecordsetHelper.ToDateTime(rs, "b2b_bsp_from_date");
                agent.IataNumber = RecordsetHelper.ToString(rs, "iata_number");
                agent.SendMailtoAllPassenger = RecordsetHelper.ToByte(rs, "send_mailto_all_passenger");
                agent.LegalName = RecordsetHelper.ToString(rs, "legal_name");
                agent.TaxId = RecordsetHelper.ToString(rs, "tax_id");
                agent.TaxIdVerifiedDateTime = RecordsetHelper.ToDateTime(rs, "tax_id_verified_date_time");
                agent.AirportTicketOfficeFlag = RecordsetHelper.ToByte(rs, "airport_ticket_office_flag");
                agent.CitySalesOfficeFlag = RecordsetHelper.ToByte(rs, "city_sales_office_flag");
                agent.UpdateBy = RecordsetHelper.ToGuid(rs, "update_by");
                agent.DefaultTicketOnPaymentFlag = RecordsetHelper.ToByte(rs, "default_ticket_on_payment_flag");
                agent.DefaultPaymentOnSaveFlag = RecordsetHelper.ToByte(rs, "default_payment_on_save_flag");
                agent.EmailInvoiceFlag = RecordsetHelper.ToByte(rs, "email_invoice_flag");
                agent.LogAvailabilityFlag = RecordsetHelper.ToByte(rs, "log_availability_flag");
                agent.ExportCycleCode = RecordsetHelper.ToString(rs, "export_cycle_code");
                agent.PosIndicator = RecordsetHelper.ToString(rs, "pos_indicator");
                agent.CashbookAgencyGroupRcd = RecordsetHelper.ToString(rs, "cashbook_agency_group_rcd");
                agent.WithholdingTaxPercentage = RecordsetHelper.ToDecimal(rs, "withholding_tax_percentage");
                agent.CommissionTopupFlag = RecordsetHelper.ToByte(rs, "commission_topup_flag");
                agent.ReceiveCommissionFlag = RecordsetHelper.ToByte(rs, "receive_commission_flag");
                agent.ConsolidatorAgencyCode = RecordsetHelper.ToString(rs, "consolidator_agency_code");
                agent.AccountingEmail = RecordsetHelper.ToString(rs, "accounting_email");
                agent.ExternalArAccount = RecordsetHelper.ToString(rs, "external_ar_account");
                agent.TaxIdVerifiedBy = RecordsetHelper.ToGuid(rs, "tax_id_verified_by");
                agent.BankIban = RecordsetHelper.ToString(rs, "bank_iban");
                agent.CommissionPercentage = RecordsetHelper.ToDecimal(rs, "commission_percentage");
                agent.CreateName = RecordsetHelper.ToString(rs, "create_name");
                agent.UpdateName = RecordsetHelper.ToString(rs, "update_name");
                agent.ProcessBaggageTagFlag = RecordsetHelper.ToByte(rs, "process_baggage_tag_flag");
                agent.ProcessRefundFlag = RecordsetHelper.ToByte(rs, "process_refund_flag");
                agent.Column1TaxRcd = RecordsetHelper.ToString(rs, "column_1_tax_rcd");
                agent.Column2TaxRcd = RecordsetHelper.ToString(rs, "column_2_tax_rcd");
                agent.Column3TaxRcd = RecordsetHelper.ToString(rs, "column_3_tax_rcd");

                agent.OwnAgencyFlag = RecordsetHelper.ToByte(rs, "own_agency_flag");
                agent.WebAgencyFlag = RecordsetHelper.ToByte(rs, "web_agency_flag");
                agent.APIFlag = RecordsetHelper.ToByte(rs, "api_flag");
                agent.B2BAllowChangeFlight = RecordsetHelper.ToByte(rs, "b2b_allow_change_flight_flag");
                agent.B2BAllowCancelFlight = RecordsetHelper.ToByte(rs, "b2b_allow_cancel_segment_flag");
                agent.B2BAllowSeat = RecordsetHelper.ToByte(rs, "b2b_allow_seat_assignment_flag");
                agent.B2BAllowServices = RecordsetHelper.ToByte(rs, "b2b_allow_service_flag");
                agent.B2BAllowPassengerInfo = RecordsetHelper.ToByte(rs, "allow_change_passenger_information_flag");
                agent.B2BAllowNameChange = RecordsetHelper.ToByte(rs, "b2b_allow_name_change_flag");
                agent.B2BAllowChangeDetail = RecordsetHelper.ToByte(rs, "b2b_allow_change_details_flag");
                //add more pad request
                agent.B2BCreditAgencyAndInvoiceFlag = RecordsetHelper.ToByte(rs, "b2b_credit_agency_and_invoice_flag");
                agent.B2BCreditCardPaymentFlag = RecordsetHelper.ToByte(rs, "b2b_credit_card_payment_flag");
                agent.B2BVoucherPaymentFlag = RecordsetHelper.ToByte(rs, "b2b_voucher_payment_flag");
                agent.B2BPostPaidFlag = RecordsetHelper.ToByte(rs, "b2b_post_paid_flag");

                agent.B2BShowRemarksFlag = RecordsetHelper.ToByte(rs, "b2b_show_remarks_flag");
                agent.B2BAllowWaitlistFlag = RecordsetHelper.ToByte(rs, "b2b_allow_waitlist_flag");
                agent.B2BAllowGroupFlag = RecordsetHelper.ToByte(rs, "b2b_allow_group_flag");
                agent.B2BAllowSplitFlag = RecordsetHelper.ToByte(rs, "b2b_allow_split_flag");

                agent.B2BGroupWaitlistFlag = RecordsetHelper.ToByte(rs, "b2b_group_waitlist_flag");
                agent.BalanceLockFlag = RecordsetHelper.ToByte(rs, "balance_lock_flag");

                agent.DisableChangesThroughB2CFlag = RecordsetHelper.ToByte(rs, "disable_changes_through_b2c_flag");
                agent.DisableWebCheckinFlag = RecordsetHelper.ToByte(rs, "disable_web_checkin_flag");
                agent.GroupFirmedFlag = RecordsetHelper.ToByte(rs, "group_firmed_flag");
                agent.GroupWaitlistFlag = RecordsetHelper.ToByte(rs, "group_waitlist_flag");
                agent.IndividualWaitlistFlag = RecordsetHelper.ToByte(rs, "individual_waitlist_flag");
                agent.AllowAddSegmentFlag = RecordsetHelper.ToByte(rs, "allow_add_segment_flag");

                agent.ChangeOfBookingAgencyCode = RecordsetHelper.ToString(rs, "change_of_booking_agency_code");
                agent.IndividualFirmedFlag = RecordsetHelper.ToByte(rs, "individual_firmed_flag");

            }
            catch
            {
                 throw;
            }
            return agent;
        }
        public static IList<User> FillUser(this IList<User> userList, ADODB.Recordset rs)
        {
            if (rs != null && rs.RecordCount > 0)
            {
                userList = new List<User>();
                User user = null;

                try
                {
                    rs.MoveFirst();
                    while (!rs.EOF)
                    {
                        user = new User();

                        user.UserAccountId = RecordsetHelper.ToGuid(rs, "user_account_id");
                        user.UserLogon = RecordsetHelper.ToString(rs, "user_logon");
                        user.UserCode = RecordsetHelper.ToString(rs, "user_code");
                        user.Lastname = RecordsetHelper.ToString(rs, "lastname");
                        user.Firstname = RecordsetHelper.ToString(rs, "firstname");
                        user.EmailAddress = RecordsetHelper.ToString(rs, "email_address");
                        user.StatusCode = RecordsetHelper.ToString(rs, "status_code");
                        user.UserPassword = RecordsetHelper.ToString(rs, "user_password");
                        user.LanguageRcd = RecordsetHelper.ToString(rs, "language_rcd");
                        user.CreateBy = RecordsetHelper.ToGuid(rs, "create_by");
                        user.CreateDateTime = RecordsetHelper.ToDateTime(rs, "create_date_time");
                        user.UpdateBy = RecordsetHelper.ToGuid(rs, "update_by");
                        user.UpdateDateTime = RecordsetHelper.ToDateTime(rs, "update_date_time");
                        user.SystemAdminFlag = RecordsetHelper.ToByte(rs, "system_admin_flag");
                        user.MakeBookingsForOthersFlag = RecordsetHelper.ToByte(rs, "make_bookings_for_others_flag");
                        user.AddressDefaultCode = RecordsetHelper.ToString(rs, "address_default_code");
                        user.ChangeSegmentFlag = RecordsetHelper.ToByte(rs, "change_segment_flag");
                        user.DeleteSegmentFlag = RecordsetHelper.ToByte(rs, "delete_segment_flag");
                        user.UpdateBookingFlag = RecordsetHelper.ToByte(rs, "update_booking_flag");
                        user.IssueTicketFlag = RecordsetHelper.ToByte(rs, "issue_ticket_flag");
                        user.CounterSalesReportFlag = RecordsetHelper.ToByte(rs, "counter_sales_report_flag");
                        user.MonFlag = RecordsetHelper.ToByte(rs, "mon_flag");
                        user.TueFlag = RecordsetHelper.ToByte(rs, "tue_flag");
                        user.WedFlag = RecordsetHelper.ToByte(rs, "wed_flag");
                        user.ThuFlag = RecordsetHelper.ToByte(rs, "thu_flag");
                        user.FriFlag = RecordsetHelper.ToByte(rs, "fri_flag");
                        user.SatFlag = RecordsetHelper.ToByte(rs, "sat_flag");
                        user.SunFlag = RecordsetHelper.ToByte(rs, "sun_flag");

                        userList.Add(user);

                        rs.MoveNext();
                    }

                }
                catch
                {
                    throw;
                }
            }
            return userList;

        }

    }
}
