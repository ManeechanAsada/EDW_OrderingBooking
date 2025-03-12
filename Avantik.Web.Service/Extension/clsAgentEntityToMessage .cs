using Avantik.Web.Service.Message.Agency;
using Avantik.Web.Service.Message.Booking;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Avantik.Web.Service.Extension
{
    public static class AgentEntityToMessage 
    {
        public static Agent ToAgentLogonMessage(this Avantik.Web.Service.Entity.Agency.Agent ag)
        {
            Agent agent = null;

            if (ag != null)
            {
                agent = new Agent();
                agent.AgencyCode = ag.AgencyCode;
                agent.CurrencyRcd = ag.CurrencyRcd;
                agent.AgencyPaymentTypeRcd = ag.AgencyPaymentTypeRcd;
                agent.AirportRcd = ag.AirportRcd;
                agent.CountryRcd = ag.CountryRcd;
                agent.LanguageRcd = ag.LanguageRcd;
                agent.AgencyPassword = ag.AgencyPassword;
                agent.DefaultUserAccountId = ag.DefaultUserAccountId;
                agent.UserLogon = ag.UserLogon;
                agent.AgencyLogon = ag.AgencyLogon;
                agent.AgencyName = ag.AgencyName;
                agent.AgLanguageRcd = ag.AgLanguageRcd;
                agent.DefaultETicketFlag = ag.DefaultETicketFlag;
                agent.Email = ag.Email;
                agent.StatusCode = ag.StatusCode;
                agent.MerchantId = ag.MerchantId;
                agent.NotifyBy = ag.NotifyBy;
                agent.DefaultCustomerDocumentId = ag.DefaultCustomerDocumentId;
                agent.DefaultSmallItineraryDocumentId = ag.DefaultSmallItineraryDocumentId;
                agent.DefaultInternalItineraryDocumentId = ag.DefaultInternalItineraryDocumentId;
                agent.PaymentDefaultCode = ag.PaymentDefaultCode;
                agent.AgencyTypeCode = ag.AgencyTypeCode;
                agent.UserAccountId = ag.UserAccountId;
                agent.UserCode = ag.UserCode;
                agent.Lastname = ag.Lastname;
                agent.Middlename = ag.Middlename;
                agent.Firstname = ag.Firstname;
                agent.OriginRcd = ag.OriginRcd;
                agent.OutstandingInvoice = ag.OutstandingInvoice;
                agent.BookingPayment = ag.BookingPayment;
                agent.AgencyAccount = ag.AgencyAccount;
                agent.WebsiteAddress = ag.WebsiteAddress;
                agent.TtyAddress = ag.TtyAddress;
                agent.CreateDateTime = ag.CreateDateTime;
                agent.UpdateDateTime = ag.UpdateDateTime;
                agent.CashbookClosingRcd = ag.CashbookClosingRcd;
                agent.CashbookClosingId = ag.CashbookClosingId;
                agent.CreateBy = ag.CreateBy;
                agent.AgencyTimezone = ag.AgencyTimezone;
                agent.SystemSettingTimezone = ag.SystemSettingTimezone;
                agent.CompanyClientProfileId = ag.CompanyClientProfileId;
                agent.ClientProfileId = ag.ClientProfileId;
                agent.InvoiceDays = ag.InvoiceDays;
                agent.AddressLine1 = ag.AddressLine1;
                agent.AddressLine2 = ag.AddressLine2;
                agent.City = ag.City;
                agent.BankCode = ag.BankCode;
                agent.BankName = ag.BankName;
                agent.BankAccount = ag.BankAccount;
                agent.ContactPerson = ag.ContactPerson;
                agent.District = ag.District;
                agent.Phone = ag.Phone;
                agent.Fax = ag.Fax;
                agent.PoBox = ag.PoBox;
                agent.Province = ag.Province;
                agent.State = ag.State;
                agent.Street = ag.Street;
                agent.ZipCode = ag.ZipCode;
                agent.B2bBspFromDate = ag.B2bBspFromDate;
                agent.IataNumber = ag.IataNumber;
                agent.SendMailtoAllPassenger = ag.SendMailtoAllPassenger;
                agent.LegalName = ag.LegalName;
                agent.TaxId = ag.TaxId;
                agent.TaxIdVerifiedDateTime = ag.TaxIdVerifiedDateTime;
                agent.AirportTicketOfficeFlag = ag.AirportTicketOfficeFlag;
                agent.CitySalesOfficeFlag = ag.CitySalesOfficeFlag;
                agent.UpdateBy = ag.UpdateBy;
                agent.DefaultTicketOnPaymentFlag = ag.DefaultTicketOnPaymentFlag;
                agent.DefaultPaymentOnSaveFlag = ag.DefaultPaymentOnSaveFlag;
                agent.EmailInvoiceFlag = ag.EmailInvoiceFlag;
                agent.LogAvailabilityFlag = ag.LogAvailabilityFlag;
                agent.ExportCycleCode = ag.ExportCycleCode;
                agent.PosIndicator = ag.PosIndicator;
                agent.CashbookAgencyGroupRcd = ag.CashbookAgencyGroupRcd;
                agent.WithholdingTaxPercentage = ag.WithholdingTaxPercentage;
                agent.CommissionTopupFlag = ag.CommissionTopupFlag;
                agent.ReceiveCommissionFlag = ag.ReceiveCommissionFlag;
                agent.ConsolidatorAgencyCode = ag.ConsolidatorAgencyCode;
                agent.AccountingEmail = ag.AccountingEmail;
                agent.ExternalArAccount = ag.ExternalArAccount;
                agent.TaxIdVerifiedBy = ag.TaxIdVerifiedBy;
                agent.BankIban = ag.BankIban;
                agent.CommissionPercentage = ag.CommissionPercentage;
                agent.CreateName = ag.CreateName;
                agent.UpdateName = ag.UpdateName;
                agent.ProcessBaggageTagFlag = ag.ProcessBaggageTagFlag;
                agent.ProcessRefundFlag = ag.ProcessRefundFlag;
                agent.Column1TaxRcd = ag.Column1TaxRcd;
                agent.Column2TaxRcd = ag.Column2TaxRcd;
                agent.Column3TaxRcd = ag.Column3TaxRcd;

                agent.OwnAgencyFlag = ag.OwnAgencyFlag;
                agent.WebAgencyFlag = ag.WebAgencyFlag;

                agent.APIFlag = ag.APIFlag;
                agent.B2BAllowCancelFlight = ag.B2BAllowCancelFlight;
                agent.B2BAllowChangeDetail = ag.B2BAllowChangeDetail;
                agent.B2BAllowChangeFlight = ag.B2BAllowChangeFlight;
                agent.B2BAllowNameChange = ag.B2BAllowNameChange;
                agent.B2BAllowPassengerInfo = ag.B2BAllowPassengerInfo;
                agent.B2BAllowSeat = ag.B2BAllowSeat;
                agent.B2BAllowServices = ag.B2BAllowServices;


                agent.B2BCreditAgencyAndInvoiceFlag = ag.B2BCreditAgencyAndInvoiceFlag;
                agent.B2BCreditCardPaymentFlag = ag.B2BCreditCardPaymentFlag;
                agent.B2BVoucherPaymentFlag = ag.B2BVoucherPaymentFlag;
                agent.B2BPostPaidFlag = ag.B2BPostPaidFlag;
                agent.B2BShowRemarksFlag = ag.B2BShowRemarksFlag;
                agent.B2BAllowWaitlistFlag = ag.B2BAllowWaitlistFlag;
                agent.B2BAllowGroupFlag = ag.B2BAllowGroupFlag;
                agent.B2BAllowSplitFlag = ag.B2BAllowSplitFlag;
                agent.B2BGroupWaitlistFlag = ag.B2BGroupWaitlistFlag;
                agent.BalanceLockFlag = ag.BalanceLockFlag;
                agent.DisableChangesThroughB2CFlag = ag.DisableChangesThroughB2CFlag;
                agent.DisableWebCheckinFlag = ag.DisableWebCheckinFlag;
                agent.GroupFirmedFlag = ag.GroupFirmedFlag;
                agent.GroupWaitlistFlag = ag.GroupWaitlistFlag;

                agent.IndividualWaitlistFlag = ag.IndividualWaitlistFlag;
                agent.AllowAddSegmentFlag = ag.AllowAddSegmentFlag;
                agent.IndividualFirmedFlag = ag.IndividualFirmedFlag;

                agent.ChangeOfBookingAgencyCode = ag.ChangeOfBookingAgencyCode;



                agent.User = ag.ToUserLogonMessage();
            }

            return agent;
        }

        public static User ToUserLogonMessage(this Avantik.Web.Service.Entity.Agency.Agent ag)
        {
            User user = null;

            if (ag != null && ag.Users != null)
            {
                user = new User();
                user.UserAccountId = ag.Users[0].UserAccountId;
                user.UserLogon = ag.Users[0].UserLogon;
                user.UserCode = ag.Users[0].UserCode;
                user.Lastname = ag.Users[0].Lastname;
                user.Firstname = ag.Users[0].Firstname;
                user.EmailAddress = ag.Users[0].EmailAddress;
                user.StatusCode = ag.Users[0].StatusCode;
                user.UserPassword = ag.Users[0].UserPassword;
                user.LanguageRcd = ag.Users[0].LanguageRcd;
                user.CreateBy = ag.Users[0].CreateBy;
                user.CreateDateTime = ag.Users[0].CreateDateTime;
                user.UpdateBy = ag.Users[0].UpdateBy;
                user.UpdateDateTime = ag.Users[0].UpdateDateTime;
                user.SystemAdminFlag = ag.Users[0].SystemAdminFlag;
                user.MakeBookingsForOthersFlag = ag.Users[0].MakeBookingsForOthersFlag;
                user.AddressDefaultCode = ag.Users[0].AddressDefaultCode;
                user.ChangeSegmentFlag = ag.Users[0].ChangeSegmentFlag;
                user.DeleteSegmentFlag = ag.Users[0].DeleteSegmentFlag;
                user.UpdateBookingFlag = ag.Users[0].UpdateBookingFlag;
                user.IssueTicketFlag = ag.Users[0].IssueTicketFlag;
                user.CounterSalesReportFlag = ag.Users[0].CounterSalesReportFlag;
                user.MonFlag = ag.Users[0].MonFlag;
                user.TueFlag = ag.Users[0].TueFlag;
                user.WedFlag = ag.Users[0].WedFlag;
                user.ThuFlag = ag.Users[0].ThuFlag;
                user.FriFlag = ag.Users[0].FriFlag;
                user.SatFlag = ag.Users[0].SatFlag;
                user.SunFlag = ag.Users[0].SunFlag;
            }

            return user;
        }

    }
}
