using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Avantik.Web.Service.Message.Agency
{
    public class Agent 
    {
        public Guid DefaultUserAccountId { get; set; }
        public Guid MerchantId { get; set; }
        public Guid DefaultCustomerDocumentId { get; set; }
        public Guid DefaultSmallItineraryDocumentId { get; set; }
        public Guid DefaultInternalItineraryDocumentId { get; set; }
        public Guid UserAccountId { get; set; }
        public Guid CashbookClosingId { get; set; }
        public Guid CreateBy { get; set; }
        public Guid CompanyClientProfileId { get; set; }
        public Guid ClientProfileId { get; set; }
        public Guid UpdateBy { get; set; }
        public Guid TaxIdVerifiedBy { get; set; }

        public byte DefaultETicketFlag { get; set; }
        public byte SendMailtoAllPassenger { get; set; }
        public byte AirportTicketOfficeFlag { get; set; }
        public byte CitySalesOfficeFlag { get; set; }
        public byte DefaultTicketOnPaymentFlag { get; set; }
        public byte DefaultPaymentOnSaveFlag { get; set; }
        public byte EmailInvoiceFlag { get; set; }
        public byte LogAvailabilityFlag { get; set; }
        public byte CommissionTopupFlag { get; set; }
        public byte ReceiveCommissionFlag { get; set; }
        public byte ProcessBaggageTagFlag { get; set; }
        public byte ProcessRefundFlag { get; set; }
        public byte WebAgencyFlag { get; set; }
        public byte OwnAgencyFlag { get; set; }

        public DateTime CreateDateTime { get; set; }
        public DateTime UpdateDateTime { get; set; }
        public DateTime B2bBspFromDate { get; set; }
        public DateTime TaxIdVerifiedDateTime { get; set; }

        public decimal OutstandingInvoice { get; set; }
        public decimal BookingPayment { get; set; }
        public decimal AgencyAccount { get; set; }
        public decimal WithholdingTaxPercentage { get; set; }
        public decimal CommissionPercentage { get; set; }

        public int AgencyTimezone { get; set; }
        public int SystemSettingTimezone { get; set; }

        public string AgencyCode { get; set; }
        public string CurrencyRcd { get; set; }
        public string AgencyPaymentTypeRcd { get; set; }
        public string AirportRcd { get; set; }
        public string CountryRcd { get; set; }
        public string LanguageRcd { get; set; }
        public string AgencyPassword { get; set; }
        public string UserLogon { get; set; }
        public string AgencyLogon { get; set; }
        public string AgencyName { get; set; }
        public string AgLanguageRcd { get; set; }
        public string Email { get; set; }
        public string StatusCode { get; set; }
        public string NotifyBy { get; set; }
        public string PaymentDefaultCode { get; set; }
        public string AgencyTypeCode { get; set; }
        public string UserCode { get; set; }
        public string Lastname { get; set; }
        public string Middlename { get; set; }
        public string Firstname { get; set; }
        public string OriginRcd { get; set; }
        public string WebsiteAddress { get; set; }
        public string TtyAddress { get; set; }
        public string CashbookClosingRcd { get; set; }
        public string InvoiceDays { get; set; }
        public string AddressLine1 { get; set; }
        public string AddressLine2 { get; set; }
        public string City { get; set; }
        public string BankCode { get; set; }
        public string BankName { get; set; }
        public string BankAccount { get; set; }
        public string ContactPerson { get; set; }
        public string District { get; set; }
        public string Phone { get; set; }
        public string Fax { get; set; }
        public string PoBox { get; set; }
        public string Province { get; set; }
        public string State { get; set; }
        public string Street { get; set; }
        public string ZipCode { get; set; }
        public string IataNumber { get; set; }
        public string LegalName { get; set; }
        public string TaxId { get; set; }
        public string ExportCycleCode { get; set; }
        public string PosIndicator { get; set; }
        public string CashbookAgencyGroupRcd { get; set; }
        public string ConsolidatorAgencyCode { get; set; }
        public string AccountingEmail { get; set; }
        public string ExternalArAccount { get; set; }
        public string BankIban { get; set; }
        public string CreateName { get; set; }
        public string UpdateName { get; set; }
        public string Column1TaxRcd { get; set; }
        public string Column2TaxRcd { get; set; }
        public string Column3TaxRcd { get; set; }
        public User User { get; set; }

        public byte APIFlag { get; set; }
        public byte B2BAllowChangeFlight { get; set; }
        public byte B2BAllowCancelFlight { get; set; }
        public byte B2BAllowSeat { get; set; }
        public byte B2BAllowServices { get; set; }
        public byte B2BAllowPassengerInfo { get; set; }
        public byte B2BAllowNameChange { get; set; }
        public byte B2BAllowChangeDetail { get; set; }


        public byte B2BCreditAgencyAndInvoiceFlag { get; set; }
        public byte B2BCreditCardPaymentFlag { get; set; }
        public byte B2BVoucherPaymentFlag { get; set; }
        public byte B2BPostPaidFlag { get; set; }
        public byte B2BShowRemarksFlag { get; set; }
        public byte B2BAllowWaitlistFlag { get; set; }
        public byte B2BAllowGroupFlag { get; set; }
        public byte B2BAllowSplitFlag { get; set; }
        public byte B2BGroupWaitlistFlag { get; set; }
        public byte BalanceLockFlag { get; set; }
        public byte DisableChangesThroughB2CFlag { get; set; }
        public byte DisableWebCheckinFlag { get; set; }
        public byte GroupFirmedFlag { get; set; }
        public byte GroupWaitlistFlag { get; set; }
        public byte IndividualWaitlistFlag { get; set; }
        public byte AllowAddSegmentFlag { get; set; }
        public byte IndividualFirmedFlag { get; set; }
        public string ChangeOfBookingAgencyCode { get; set; }

    }
}
