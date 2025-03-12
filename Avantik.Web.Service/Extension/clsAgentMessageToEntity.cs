using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Avantik.Web.Service.Entity.Agency;

namespace Avantik.Web.Service.Extension
{
    public static class AgentMessageToEntity
    {
        public static Agent ToAgentEntity(this Message.Agency.Agent ag)
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

            }
            return agent;
        }
    }
}
