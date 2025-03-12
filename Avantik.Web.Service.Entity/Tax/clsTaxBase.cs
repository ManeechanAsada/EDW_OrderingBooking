using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Avantik.Web.Service.Entity.Tax
{
    public abstract class TaxBase
    {
        public string tax_rcd { get; set; }
        public string airport_rcd { get; set; }
        public string country_rcd { get; set; }

        public byte transit_tax_only_flag { get; set; }
        public byte include_surcharge_flag { get; set; }
        public byte minimum_tax_amount_flag { get; set; }
        public byte fare_basis_flag { get; set; }
    }
}
