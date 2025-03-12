using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ServiceModel;
namespace Avantik.Web.Service.Message.OrderBooking
{
    [MessageContract]
    public class Remark
    {
        [MessageBodyMember]
        public DateTime timelimit_date_time { get; set; }
        [MessageBodyMember]
        public DateTime utc_timelimit_date_time { get; set; }
        [MessageBodyMember]
        public string remark_type_rcd { get; set; }
        [MessageBodyMember]
        public string remark_text { get; set; }
        [MessageBodyMember]
        public string nickname { get; set; }
    }
}
