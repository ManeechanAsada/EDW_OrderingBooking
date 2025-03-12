using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

namespace Avantik.Web.Service.Entity
{
    public class Country 
    {
        #region Property

        public string CountryRcd  { get; set; }
        public string DisplayName  { get; set; }
        public string StatusCode  { get; set; }
        public string CurrencyRcd  { get; set; }
        public string VatDisplay  { get; set; }
        public string CountryCodeLong  { get; set; }
        public string CountryNumber  { get; set; }
        public string PhonePrefix  { get; set; }

        public byte IssueCountryFlag  { get; set; }
        public byte ResidenceCountryFlag  { get; set; }
        public byte NationalityCountryFlag  { get; set; }
        public byte AddressCountryFlag  { get; set; }
        public byte AgencyCountryFlag { get; set; }

        #endregion
    }
}
