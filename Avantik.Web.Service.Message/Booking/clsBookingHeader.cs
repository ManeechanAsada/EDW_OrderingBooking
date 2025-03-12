using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.Serialization;
using System.ServiceModel;

namespace Avantik.Web.Service.Message.Booking
{
    public class BookingHeader
    {
        [MessageBodyMember]
        public Guid BookingId { get; set; }
        [MessageBodyMember]
        public string BookingSourceRcd { get; set; }
        [MessageBodyMember]
        public string CurrencyRcd { get; set; }
        [MessageBodyMember]
        public Guid ClientProfileId { get; set; }
        [MessageBodyMember]
        public long BookingNumber { get; set; }
        [MessageBodyMember]
        public string RecordLocator { get; set; }
        [MessageBodyMember]
        public int NumberOfAdults { get; set; }
        [MessageBodyMember]
        public int NumberOfChildren { get; set; }
        [MessageBodyMember]
        public int NumberOfInfants { get; set; }
        [MessageBodyMember]
        public string LanguageRcd { get; set; }
        [MessageBodyMember]
        public string AgencyCode { get; set; }
        [MessageBodyMember]
        public string ContactName { get; set; }
        [MessageBodyMember]
        public string ContactEmail { get; set; }
        [MessageBodyMember]
        public string PhoneMobile { get; set; }
        [MessageBodyMember]
        public string PhoneHome { get; set; }
        [MessageBodyMember]
        public string PhoneBusiness { get; set; }
        [MessageBodyMember]
        public string ReceivedFrom { get; set; }
        [MessageBodyMember]
        public string PhoneFax { get; set; }
        [MessageBodyMember]
        public string PhoneSearch { get; set; }
        [MessageBodyMember]
        public string Comment { get; set; }
        [MessageBodyMember]
        public byte NotifyByEmailFlag { get; set; }
        [MessageBodyMember]
        public byte NotifyBySmsFlag { get; set; }
        [MessageBodyMember]
        public string GroupName { get; set; }
        [MessageBodyMember]
        public byte GroupBookingFlag { get; set; }
        [MessageBodyMember]
        public string AgencyName { get; set; }
        [MessageBodyMember]
        public byte OwnAgencyFlag { get; set; }
        [MessageBodyMember]
        public byte WebAgencyFlag { get; set; }
        [MessageBodyMember]
        public long ClientNumber { get; set; }
        [MessageBodyMember]
        public string Lastname { get; set; }
        [MessageBodyMember]
        public string Firstname { get; set; }
        [MessageBodyMember]
        public string City { get; set; }
        [MessageBodyMember]
        public string CreateName { get; set; }
        [MessageBodyMember]
        public string Street { get; set; }
        [DataMember(EmitDefaultValue = false)]
        public DateTime LockDateTime { get; set; }
        [MessageBodyMember]
        public string Middlename { get; set; }
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
        public string TitleRcd { get; set; }
        [MessageBodyMember]
        public string ExternalPaymentReference { get; set; }
        [MessageBodyMember]
        public Guid CreateBy { get; set; }
        [MessageBodyMember]
        public Guid UpdateBy { get; set; }
        [DataMember(EmitDefaultValue = false)]
        public DateTime CreateDateTime { get; set; }
        [DataMember(EmitDefaultValue = false)]
        public DateTime UpdateDateTime { get; set; }
        [MessageBodyMember]
        public string CostCenter { get; set; }
        [MessageBodyMember]
        public string PurchaseOrder { get; set; }
        [MessageBodyMember]
        public string ProjectNumber { get; set; }
        [MessageBodyMember]
        public string LockName { get; set; }
        [MessageBodyMember]
        public string IpAddress { get; set; }
        [MessageBodyMember]
        public byte ApprovalFlag { get; set; }
        [MessageBodyMember]
        public string InvoiceReceiver { get; set; }
        [MessageBodyMember]
        public string TaxId { get; set; }
        [MessageBodyMember]
        public byte NewsletterFlag { get; set; }
        [MessageBodyMember]
        public string ContactEmailCc { get; set; }
        [MessageBodyMember]
        public string MobileEmail { get; set; }
        [MessageBodyMember]
        public string VendorRcd { get; set; }
        [DataMember(EmitDefaultValue = false)]
        public DateTime BookingDateTime { get; set; }
        [MessageBodyMember]
        public byte NoVatFlag { get; set; }
        [MessageBodyMember]
        public string CompanyName { get; set; }
        [MessageBodyMember]
        public byte BusinessFlag { get; set; }

    }
}
