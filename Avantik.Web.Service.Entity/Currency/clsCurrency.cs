using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

namespace Avantik.Web.Service.Entity
{
    public class Currency 
    {
        #region Property

        public string CurrencyRcd { get; set; }
        public string DisplayName { get; set; }
        public string CurrencyNumber { get; set; }
        public string DisplayCode { get; set; }

        public decimal MaxVoucherValue { get; set; }
        public decimal RoundingRule { get; set; }

        public Int16 NumberOfDecimals { get; set; }

        #endregion
    }
}
