using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Avantik.Web.Service.Entity.Booking
{
    public class BookingHeader
    {
        public Guid BookingId { get; set; }
        public string BookingSourceRcd { get; set; }
        public string CurrencyRcd { get; set; }
        public Guid ClientProfileId { get; set; }
        public long BookingNumber { get; set; }
        public string RecordLocator { get; set; }
        public int NumberOfAdults { get; set; }
        public int NumberOfChildren { get; set; }
        public int NumberOfInfants { get; set; }
        public string LanguageRcd { get; set; }
        public string AgencyCode { get; set; }
        public string ContactName { get; set; }
        public string ContactEmail { get; set; }
        public string PhoneMobile { get; set; }
        public string PhoneHome { get; set; }
        public string PhoneBusiness { get; set; }
        public string ReceivedFrom { get; set; }
        public string PhoneFax { get; set; }
        public string PhoneSearch { get; set; }
        public string Comment { get; set; }
        public byte NotifyByEmailFlag { get; set; }
        public byte NotifyBySmsFlag { get; set; }
        public string GroupName { get; set; }
        public byte GroupBookingFlag { get; set; }
        public string AgencyName { get; set; }
        public byte OwnAgencyFlag { get; set; }
        public byte WebAgencyFlag { get; set; }
        public long ClientNumber { get; set; }
        public string Lastname { get; set; }
        public string Firstname { get; set; }
        public string City { get; set; }
        public string CreateName { get; set; }
        public string Street { get; set; }
        public DateTime LockDateTime { get; set; }
        public string Middlename { get; set; }
        public string AddressLine1 { get; set; }
        public string AddressLine2 { get; set; }
        public string State { get; set; }
        public string District { get; set; }
        public string Province { get; set; }
        public string ZipCode { get; set; }
        public string PoBox { get; set; }
        public string CountryRcd { get; set; }
        public string TitleRcd { get; set; }
        public string ExternalPaymentReference { get; set; }
        public Guid CreateBy { get; set; }
        public Guid UpdateBy { get; set; }
        public DateTime CreateDateTime { get; set; }
        public DateTime UpdateDateTime { get; set; }
        public string CostCenter { get; set; }
        public string PurchaseOrder { get; set; }
        public string ProjectNumber { get; set; }
        public string LockName { get; set; }
        public string IpAddress { get; set; }
        public byte ApprovalFlag { get; set; }
        public string InvoiceReceiver { get; set; }
        public string TaxId { get; set; }
        public byte NewsletterFlag { get; set; }
        public string ContactEmailCc { get; set; }
        public string MobileEmail { get; set; }
        public string VendorRcd { get; set; }
        public DateTime BookingDateTime { get; set; }
        public byte NoVatFlag { get; set; }
        public string CompanyName { get; set; }
        public byte BusinessFlag { get; set; }

    }
}
