using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.Serialization;
using System.ServiceModel;

namespace Avantik.Web.Service.Message.Fee
{
    public class ServiceFee : SegmentService
    {
        [MessageBodyMember]
        public string FeeRcd { get; set; }
        [MessageBodyMember]
        public string DisplayName { get; set; }
        [MessageBodyMember]
        public string CurrencyRcd { get; set; }
        [MessageBodyMember]
        public decimal FeeAmount { get; set; }
        [MessageBodyMember]
        public decimal FeeAmountIncl { get; set; }
        [MessageBodyMember]
        public decimal TotalFeeAmount { get; set; }
        [MessageBodyMember]
        public decimal TotalFeeAmountIncl { get; set; }
        [MessageBodyMember]
        public bool ServiceOnRequestFlag { get; set; }
        [MessageBodyMember]
        public bool CutOffTime { get; set; }

    }
}
