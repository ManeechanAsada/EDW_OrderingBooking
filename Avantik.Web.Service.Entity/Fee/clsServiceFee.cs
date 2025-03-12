using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

namespace Avantik.Web.Service.Entity
{
    public class ServiceFee : SegmentService
    {
        public string FeeRcd { get; set; }
        public string DisplayName { get; set; }
        public string CurrencyRcd { get; set; }
        public decimal FeeAmount { get; set; }
        public decimal FeeAmountIncl { get; set; }
        public decimal TotalFeeAmount { get; set; }
        public decimal TotalFeeAmountIncl { get; set; }
        public bool ServiceOnRequestFlag { get; set; }
        public bool CutOffTime { get; set; }

    }
}
