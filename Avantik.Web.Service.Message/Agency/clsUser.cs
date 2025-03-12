using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Avantik.Web.Service.Message.Agency
{
    public class User
    {
        public Guid UserAccountId { get; set; }
        public Guid CreateBy { get; set; }
        public Guid UpdateBy { get; set; }

        public string UserLogon { get; set; }
        public string UserCode { get; set; }
        public string Lastname { get; set; }
        public string Firstname { get; set; }
        public string EmailAddress { get; set; }
        public string StatusCode { get; set; }
        public string UserPassword { get; set; }
        public string LanguageRcd { get; set; }
        public string AddressDefaultCode { get; set; }

        public DateTime CreateDateTime { get; set; }
        public DateTime UpdateDateTime { get; set; }

        public byte SystemAdminFlag { get; set; }
        public byte MakeBookingsForOthersFlag { get; set; }
        public byte ChangeSegmentFlag { get; set; }
        public byte DeleteSegmentFlag { get; set; }
        public byte UpdateBookingFlag { get; set; }
        public byte IssueTicketFlag { get; set; }
        public byte CounterSalesReportFlag { get; set; }
        public byte MonFlag { get; set; }
        public byte TueFlag { get; set; }
        public byte WedFlag { get; set; }
        public byte ThuFlag { get; set; }
        public byte FriFlag { get; set; }
        public byte SatFlag { get; set; }
        public byte SunFlag { get; set; }

    }
}
