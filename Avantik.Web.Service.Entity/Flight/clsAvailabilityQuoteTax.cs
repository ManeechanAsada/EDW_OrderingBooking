using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

namespace Avantik.Web.Service.Entity.Flight
{
    public class AvailabilityQuoteTax : Tax.TaxBase
    {
        public string valid_for_class { get; set; }
        
        public string fare_code { get; set; }
        public string tax_currency { get; set; }
        
        public decimal tax_amount { get; set; }
        public decimal vat_percentage { get; set; }
        public decimal tax_percentage { get; set; }
        public decimal tax_amount_incl { get; set; }
        public decimal exchange_to_accounting { get; set; }
        public decimal exchange_from_accounting { get; set; }
        public decimal origin_tax_distribution { get; set; }
        public decimal destination_tax_distribution { get; set; }

        public DateTime departure_date { get; set; }
        public string passenger_type_rcd { get; set; }
    }
}
